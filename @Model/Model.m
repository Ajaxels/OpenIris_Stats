classdef Model < handle
    % classdef Model < handle
    % the main model class of OpenIrisStats
    
    % Copyright (C) 2022 Ilya Belevich, University of Helsinki (ilya.belevich @ helsinki.fi)
    % The MIT License (https://opensource.org/licenses/MIT)
    
    properties 
        T
        % table with imported excel sheet
        Settings
        % a structure with settings
        VariableNames
        % list of column names from excel sheet
    end
    
    properties (SetObservable)

    end
        
    events
        updateGuiWidgets
        % event after update of GuiWidgets of Controller
    end
    
    methods
        % declaration of functions in the external files, keep empty line in between for the doc generator
        
        BatchOptOut = selectFile(obj, BatchOpt)   % choose a file
        
        function obj = Model()
            obj.reset();
        end
        
        function reset(obj)
            obj.T = [];  
            obj.Settings = struct();     % current session settings
            obj.Settings.gui = struct();     % current settings for gui widgets
            obj.Settings.gui.ListOfInvoices = [];   % full filename list with invoices
            obj.Settings.gui.OutputFilename = fullfile(pwd, 'openiris_stats.xls');
            obj.Settings.gui.PriceTypeIndexField = 'PriceType';    % name of the field with index of the price type
            obj.Settings.gui.HeaderStartingCell = 'A1';
            obj.Settings.gui.DataStartingCell = 'A2';
        end
        
        function getColumnNames(obj)
            % get column names from the excel file
            wb = waitbar(0, sprintf('Obtaining names for the column\nPlease wait...'));
            warning('off', 'MATLAB:table:ModifiedAndSavedVarnames');
            waitbar(0.05, wb);
            rangeText = sprintf('%s:%s', obj.Settings.gui.HeaderStartingCell(2:end), obj.Settings.gui.DataStartingCell(2:end));
            obj.T = readtable(obj.Settings.gui.ListOfInvoices{1}, 'Range', rangeText, 'UseExcel', false);   % read excel file
            waitbar(0.9, wb);
            obj.VariableNames = obj.T.Properties.VariableNames;
            waitbar(1, wb);
            delete(wb);
        end
        
        function start(obj)
            tic
            warning('off', 'MATLAB:xlswrite:AddSheet');
            wb = waitbar(0, sprintf('Calculating statistics\nPlease wait...'), 'Name', 'OpenIris Stats');
            
            for fnId = 1:numel(obj.Settings.gui.ListOfInvoices)
                fnIn = obj.Settings.gui.ListOfInvoices{fnId};
                opts = detectImportOptions(fnIn, 'NumHeaderLines', 0);
                opts.VariableNamesRange = obj.Settings.gui.HeaderStartingCell;
                opts.DataRange = obj.Settings.gui.DataStartingCell;
                opts = setvartype(opts, 'char');    % force to return elements as cell array
                if fnId == 1
                    obj.T = readtable(fnIn, opts, 'ReadVariableNames', true, 'UseExcel', false);   % read excel file
                    obj.VariableNames = obj.T.Properties.VariableNames;

                    GroupPIIndex = find(ismember(obj.VariableNames, 'Group'));
                    ProjectNameIndex = find(ismember(obj.VariableNames, 'RequestTitle'));
                    RequestIDIndex = find(ismember(obj.VariableNames, 'RequestID'));
                    OrganizationIndex = find(ismember(obj.VariableNames, 'Organization'));  
                    PriceTypeIndex = find(ismember(obj.VariableNames, 'PriceType')); 
                    ChargeIndex = find(ismember(obj.VariableNames, 'Charge'));
                    ResourceIndex = find(ismember(obj.VariableNames, 'Resource'));
                    ChargeTypeIndex = find(ismember(obj.VariableNames, 'ChargeType'));
                    UserNameIndex = find(ismember(obj.VariableNames, 'UserName'));
                    UserEMailNameIndex = find(ismember(obj.VariableNames, 'UserEmail'));
                    DescriptionIndex = find(ismember(obj.VariableNames, 'Description'));
                    BookingStartIndex = find(ismember(obj.VariableNames, 'BookingStart'));
                    BookingEndIndex = find(ismember(obj.VariableNames, 'BookingEnd'));
                else
                    try
                        dummyT = readtable(fnIn, opts, 'ReadVariableNames', true, 'UseExcel', false);   % read excel file
                        dummyVariableNames = dummyT.Properties.VariableNames;
                        % find potential new columns
                        additionalFieldsIndices = find(~ismember(dummyVariableNames, obj.VariableNames));
                        if ~isempty(additionalFieldsIndices)
                            % add additional columns to the main table
                            for extraFieldId = 1:numel(additionalFieldsIndices)
                                obj.T.(dummyVariableNames{additionalFieldsIndices(extraFieldId)}) = cell([size(obj.T,1), 1]);
                            end
                            obj.VariableNames = obj.T.Properties.VariableNames;
                        end
                        obj.T = [obj.T; dummyT];
                    catch err
                        errordlg(sprintf('Processing:\n%s\n\n%s', fnIn, err.message), 'Importing error');
                        delete(wb);
                        return;
                    end
                end
                waitbar(fnId/numel(obj.Settings.gui.ListOfInvoices), wb);
            end

            %obj.Settings.gui.StatsNumberOfProjects = true;
            %obj.Settings.gui.StatsNumberOfUsers = true;
            %obj.Settings.gui.StatsNumberOfGroups = true;
            %obj.Settings.gui.StatsNumberOfMicroscopeHours = true;
            %obj.Settings.gui.SplitNumbers = true;

            Summary = cell([50, 60]);     % allocate space
            Summary{2,1} = 'Parameter';

            % get number of price groups and add one extra for totals
            % Default comes from Products that do not have separate
            % categories for different customers
            priceCategories = table2cell(unique(obj.T(:, PriceTypeIndex)));
            priceCategories = [priceCategories; {'Total'}]';
            noPriceGroups = numel(priceCategories);
            %                         {'Default'             }
            %                         {'HiLIFE Commercial'   }
            %                         {'HiLIFE External'     }
            %                         {'HiLIFE internal'     }
            %                         {'HiLife International'}
            %                         {'Total'}
            Summary(2,2:2+noPriceGroups-1) = priceCategories;
            waitbar(0.1, wb);

            % ------------------------------------------------------------
            % get statistics for number of projects
            rowId = 3;
            Summary(rowId, 1) = {'Number of projects'}; 
            [noPriceGroupsCounts, defaultPriceTypeIds, projectHitsTable] = obj.countCategories(RequestIDIndex, PriceTypeIndex, priceCategories);
            if isempty(noPriceGroupsCounts); delete(wb); return; end
            Summary(rowId, 2:2+noPriceGroups-1) = num2cell(noPriceGroupsCounts);
            if ~isempty(defaultPriceTypeIds)
                Summary(1, 2+noPriceGroups+1) = {'Identifiers for Default price type'};
                Summary(rowId, 2+noPriceGroups+1:2+noPriceGroups+1+numel(defaultPriceTypeIds)-1) = defaultPriceTypeIds;
            end
            rowId = rowId + 1;
            waitbar(0.2, wb);
            
            % ------------------------------------------------------------
            % get statistics for number of groups
            Summary(rowId, 1) = {'Number of user groups'};
            [noPriceGroupsCounts, defaultPriceTypeIds, groupsHitsTable] = obj.countCategories(GroupPIIndex, PriceTypeIndex, priceCategories);
            if isempty(noPriceGroupsCounts); delete(wb); return; end
            Summary(rowId,2:2+noPriceGroups-1) = num2cell(noPriceGroupsCounts);
            if ~isempty(defaultPriceTypeIds)
                Summary(rowId, 2+noPriceGroups+1:2+noPriceGroups+1+numel(defaultPriceTypeIds)-1) = defaultPriceTypeIds;
            end
            rowId = rowId + 1;
            waitbar(0.3, wb);

            % ------------------------------------------------------------
            % get statistics for number of users
            Summary(rowId, 1) = {'Number of users'};
            [noPriceGroupsCounts, defaultPriceTypeIds, usersHitsTable] = obj.countCategories(UserNameIndex, PriceTypeIndex, priceCategories);
            [~, ~, usersEmailHitsTable] = obj.countCategories(UserEMailNameIndex, PriceTypeIndex, priceCategories);
            if isempty(noPriceGroupsCounts); delete(wb); return; end
            Summary(rowId,2:2+noPriceGroups-1) = num2cell(noPriceGroupsCounts);
            if ~isempty(defaultPriceTypeIds)
                Summary(rowId, 2+noPriceGroups+1:2+noPriceGroups+1+numel(defaultPriceTypeIds)-1) = defaultPriceTypeIds;
            end
            rowId = rowId + 1;
            waitbar(0.4, wb);
            
            % -------------------------------------------------------
            % get stats for the instrument hours
            % get indices of the products
            productIndices = ismember(table2array(obj.T(:, ChargeTypeIndex)), 'Product (request)');
            % generate table with reservations
            reservationsTable = obj.T(~productIndices, :);
            
            if isa(reservationsTable.BookingStart(1), 'datetime') % the dates in charges are imported already as datetime class
                startingDates = reservationsTable.BookingStart;
                endingDates = reservationsTable.BookingEnd;
            else
                startingDates = datetime(table2cell(reservationsTable(:, BookingStartIndex)), 'InputFormat','yy-MM-dd HH:mm');
                endingDates = datetime(table2cell(reservationsTable(:, BookingEndIndex)), 'InputFormat','yy-MM-dd HH:mm');
            end
            timesVec = endingDates-startingDates;
            
            [splitEntries, ~, ic] = unique(reservationsTable(:, PriceTypeIndex));
            splitEntries = table2cell(splitEntries);
            noPriceGroupsCounts = cell([1, numel(priceCategories)]);
            totalHours = duration(0, 0, 0);
            for entryId = 1:numel(splitEntries)
                vals = timesVec(find(ismember(ic, entryId)));
                sumTime = sum(vals);
                noPriceGroupsCounts{ismember(priceCategories, splitEntries{entryId})} = sumTime;
                totalHours = totalHours + sumTime;
            end
            noPriceGroupsCounts{end} = totalHours;
            Summary(rowId, 1) = {'Usage hours:'};
            Summary(rowId, 2:2+noPriceGroups-1) = noPriceGroupsCounts;
            rowId = rowId + 1;
            waitbar(0.5, wb);

            % -------------------------------------------------
            % calculate number of trainings
            trainingIndices = ismember(table2array(reservationsTable(:, ChargeTypeIndex)), 'Scheduled (Training)');
            trainingTable = reservationsTable(trainingIndices, :);
            
            [noPriceGroupsCounts, defaultPriceTypeIds, trainingHitsTable] = obj.countCategories(UserNameIndex, PriceTypeIndex, priceCategories, trainingTable);
            if isempty(noPriceGroupsCounts); delete(wb); return; end
            Summary(rowId, 1) = {'Trainings:'};
            Summary(rowId, 2:2+noPriceGroups-1) = num2cell(noPriceGroupsCounts);
            rowId = rowId + 1;
            waitbar(0.6, wb);

            % -----------------------------------------------
            % save Excel
            if exist(obj.Settings.gui.OutputFilename, 'file') == 2; delete(obj.Settings.gui.OutputFilename); end
            writecell(Summary, obj.Settings.gui.OutputFilename, 'Sheet', 'Stats');

            % write additional sheets
            writetable(projectHitsTable, obj.Settings.gui.OutputFilename, 'Sheet', 'Projects');
            writetable(groupsHitsTable, obj.Settings.gui.OutputFilename, 'Sheet', 'Groups');
            writetable(usersHitsTable, obj.Settings.gui.OutputFilename, 'Sheet', 'Users');
            writetable(usersEmailHitsTable, obj.Settings.gui.OutputFilename, 'Sheet', 'UsersEmails');
            writetable(trainingHitsTable, obj.Settings.gui.OutputFilename, 'Sheet', 'Trainings');
            
            waitbar(1, wb);
            delete(wb);

            return;

            % remove default sheets 1,2,3
            objExcel = actxserver('Excel.Application');
            objExcel.Workbooks.Open(fullfile(obj.Settings.gui.OutputFilename)); % Full path is necessary!
            % Delete sheets 1, 2, 3.
            try
                objExcel.ActiveWorkbook.Worksheets.Item(1).Delete;
                objExcel.ActiveWorkbook.Worksheets.Item(1).Delete;
                objExcel.ActiveWorkbook.Worksheets.Item(1).Delete;
            catch err
                err;
            end
            sheet = objExcel.ActiveWorkbook.Sheets.Item(1);
            sheet.Activate;
            objExcel.ActiveSheet.Range('C1').Font.Bold = true;
            objExcel.ActiveSheet.Range('C1').Font.Size = 16;
                
            objExcel.ActiveSheet.Range('C13:F13').Font.Bold = true;
            objExcel.ActiveSheet.Range('C16:H16').Font.Bold = true;
            objExcel.ActiveSheet.Range('D3').Font.Bold = true;  % Billing address
                
            rangeText = sprintf('A1:A%d', shiftY);
            objExcel.ActiveSheet.Range(rangeText).Font.Bold = true;
            objExcel.ActiveSheet.Range(rangeText).HorizontalAlignment = 4;  % 2-left, 3-center, 4 - right
                
            objExcel.ActiveSheet.Columns.Item(1).columnWidth = 32; % 1st column width
            objExcel.ActiveSheet.Columns.Item(2).columnWidth = 1; % 2nd column width
            objExcel.ActiveSheet.Columns.Item(3).columnWidth = 35;
            objExcel.ActiveSheet.Columns.Item(4).columnWidth = 22;
            objExcel.ActiveSheet.Columns.Item(5).columnWidth = 18;
            objExcel.ActiveSheet.Columns.Item(6).columnWidth = 18;
            objExcel.ActiveSheet.Columns.Item(7).columnWidth = 18;
            objExcel.ActiveSheet.Columns.Item(8).columnWidth = 18;
            objExcel.ActiveSheet.Columns.Item(9).columnWidth = 18;

            % merge billing address cells
            objExcel.ActiveSheet.Range('E3:F11').MergeCells = 1;
            objExcel.ActiveSheet.Range('E3').VerticalAlignment = -4160; % align to top, see more https://docs.microsoft.com/en-us/office/vba/api/excel.xlvalign
            objExcel.ActiveSheet.Range('E3').WrapText = 1;
            objExcel.ActiveSheet.Range('C8:D9').MergeCells = 1;
            objExcel.ActiveSheet.Range('C8').VerticalAlignment = -4160;
            objExcel.ActiveSheet.Range('C8').WrapText = 1;

            % add borders
            objExcel.ActiveSheet.Range('A1:F1').Borders.Item('xlEdgeBottom').LineStyle = 1;
            for i=1:numel(lineVec)
                rangeText = sprintf('A%d:F%d', lineVec(i), lineVec(i));
                objExcel.ActiveSheet.Range(rangeText).Borders.Item('xlEdgeBottom').LineStyle = 1;
                objExcel.ActiveSheet.Range(rangeText).Font.Bold = true;
                rangeText = sprintf('A%d', lineVec(i));
                objExcel.ActiveSheet.Range(rangeText).Font.Size = 12;
            end

            % highlight the summary section
            rangeText = sprintf('A%d:F%d', lineVec(end), lineVec(end));
            objExcel.ActiveSheet.Range(rangeText).Interior.ColorIndex = 19; %40; % RGB(r, g, b)
            rangeText = sprintf('F%d', lineVec(end)+1);
            objExcel.ActiveSheet.Range(rangeText).Interior.ColorIndex = 40; %45; % RGB(r, g, b)
            objExcel.ActiveSheet.Range(rangeText).Font.Bold = true;

            % add alignment
            rangeText = sprintf('D8:H%d', shiftY);
            objExcel.ActiveSheet.Range(rangeText).HorizontalAlignment = 3;  % 2-left, 3-center, 4 - right

            objExcel.PrintCommunication = 1;
            try     % when no printers installed the code below gives an error
                objExcel.ActiveSheet.PageSetup.Zoom = false;
                objExcel.ActiveSheet.PageSetup.FitToPagesWide = 1;
                %objExcel.ActiveSheet.PageSetup.PrintArea = sprintf('A1:F%d', size(s, 1));  % "$A$1:$C$5";
            catch err

            end
            % Save, close and clean up.
            objExcel.ActiveWorkbook.Save;
            objExcel.ActiveWorkbook.Close;
            objExcel.Quit;
            objExcel.delete;
            waitbar(1, wb);

            delete(wb);
            toc
        end

        function [noPriceGroupsCounts, defaultPriceTypeIds, hitsTable] = countCategories(obj, splitIndex, priceTypeIndex, priceCategories, tableIn)
            % function [noPriceGroupsCounts, defaultPriceTypeIds, hitsTable] = countCategories(obj, splitIndex, priceTypeIndex, priceCategories, tableIn);
            % calculate statstics for the selected category
            %
            % Parameters:
            % splitIndex: index of the column for calculation of statistics
            % priceTypeIndex: index of the column with the price type
            % information for splitting the values in "splitIndex" into subcategoires
            % priceCategories: cell array, with all possible categories,
            % for example: {'Default'}    {'HiLIFE Commercial'}    {'HiLIFE External'}    {'HiLIFE internal'}    {'HiLife Internat…'}    {'Total'}
            % ending with Total.
            % tableIn: [optional], table to process, when missing will process obj.T
            %
            % Return values:
            % noPriceGroupsCounts: vector with counts for each category
            % defaultPriceTypeIds: list of project Ids that have only Default price tag
            % hitsTable: a table with the detected unique hits

            if nargin < 5; tableIn = obj.T; end
            defaultPriceTypeIds = [];   % list of project numbers that have only Default type

            noPriceGroupsCounts = zeros([1, numel(priceCategories)]);
            
            % simplify the tableIn and extract unique values
            splitEntries = table2cell(unique(tableIn(:, [splitIndex priceTypeIndex])));
            %  {'10740'}    {'HiLIFE internal'     }
            %  {'10760'}    {'Default'             }
            %  {'10760'}    {'HiLIFE internal'     }
            %  {'10761'}    {'Default'             }
            %  {'10761'}    {'HiLIFE internal'     }

            % allocate space for the list output of hits
            hitsTable = table('Size',[size(splitEntries,1), numel(priceCategories)],'VariableTypes', repmat({'string'}, [1, numel(priceCategories)]), 'VariableNames', priceCategories);

            % check for empty fields
            missingSplitFieldValues = find(cellfun(@isempty, splitEntries));
            if ~isempty(missingSplitFieldValues)
                warndlg(sprintf('!!! Warning !!!\n\nPlease check the "%s" field in the original Excel file!\nIt looks that some of those fields are empty.\nTo proceed further all fields used for splitting should contain some data!', strjoin(obj.VariableNames([splitIndex priceTypeIndex]), ', ')), ...
                    'Missing info');
                noPriceGroupsCounts = [];
            end

            if ~isa(splitEntries{1,1}, 'char')  % convert to char if the request number is fetched as double
                splitEntries(:,1) = cellfun(@num2str, splitEntries(:,1), 'UniformOutput', false);
            end
            [C, ~, ic] = unique(splitEntries(:,1));
            
            for entryId = 1:numel(C)
                vals = splitEntries(find(ismember(ic, entryId)),:);
                if size(vals, 1) == 1
                    if strcmp(vals{1,2}, 'Default')
                        defaultPriceTypeIds = [defaultPriceTypeIds, vals(1, 1)];
                    end
                    % noPriceGroupsCounts  = noPriceGroupsCounts + ismember(priceCategories, vals);
                    catIndex = find(ismember(priceCategories, vals));
                    noPriceGroupsCounts(catIndex) = noPriceGroupsCounts(catIndex) + 1;
                    hitsTable.(priceCategories{catIndex})(noPriceGroupsCounts(catIndex)) = vals(1);
                else
                    notOk = true;
                    i = 1;
                    while notOk
                        if ~strcmp(vals{i,2}, 'Default')
                            % noPriceGroupsCounts  = noPriceGroupsCounts + ismember(priceCategories, vals(i, 2));
                            catIndex = find(ismember(priceCategories, vals(i, 2)));
                            noPriceGroupsCounts(catIndex) = noPriceGroupsCounts(catIndex) + 1;
                            hitsTable.(priceCategories{catIndex})(noPriceGroupsCounts(catIndex)) = vals(i, 1);
                            notOk = false;
                        else
                            if i == size(vals,1) % reach end of the list, add 'Default'
                                %noPriceGroupsCounts  = noPriceGroupsCounts + ismember(priceCategories, vals(i, 2));
                                catIndex = find(ismember(priceCategories, vals(i, 2)));
                                noPriceGroupsCounts(catIndex) = noPriceGroupsCounts(catIndex) + 1;
                                hitsTable.(priceCategories{catIndex})(noPriceGroupsCounts(catIndex)) = vals(i, 1);

                                notOk = false;
                                defaultPriceTypeIds = [defaultPriceTypeIds, vals(i, 1)];
                            end
                            i = i + 1;
                        end
                    end
                end
                noPriceGroupsCounts(end)  = noPriceGroupsCounts(end) + 1; %  totals
            end
            %noPriceGroupsCounts(end)  = numel(C); % totals
            % clip the missing values
            hitsTable = hitsTable(1:max(noPriceGroupsCounts(1:end-1)),:);
        end
    end

end