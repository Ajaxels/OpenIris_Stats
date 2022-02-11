classdef Controller < handle
    % @type Controller class is a template class for using with
    % GUI developed using appdesigner of Matlab
    
	% Copyright (C) 2022 Ilya Belevich, University of Helsinki (ilya.belevich @ helsinki.fi)
    % The MIT License (https://opensource.org/licenses/MIT)
    
    
    properties
        Model
        % handles to the model
        View
        % handle to the view
        listener
        % a cell array with handles to listeners
        childControllers
        % list of opened subcontrollers
        childControllersIds
        % a cell array with names of initialized child controllers
    end
    
    events
        %> Description of events
        closeEvent
        % event firing when window is closed
    end
    
    methods (Static)
        function purgeControllers(obj, src, evnt)
            % find index of the child controller
            id = obj.findChildId(class(src));
            
            % delete the child controller
            delete(obj.childControllers{id});
            
            % clear the handle
            obj.childControllers(id) = [];
            obj.childControllersIds(id) = [];
        end
        
        
        function ViewListner_Callback(obj, src, evnt)
            switch evnt.EventName
                case {'updateGuiWidgets'}
                    obj.updateWidgets();
            end
        end
    end
    
    methods
        % declaration of functions in the external files, keep empty line in between for the doc generator
        id = findChildId(obj, childName)        % find index of a child controller  
        
        startController(obj, controllerName, varargin)        % start a child controller
        
        function obj = Controller(Model, parameter)
            obj.Model = Model;    % assign model
            obj.View = View(obj);
            
            obj.View.gui.Name = [obj.View.gui.Name, '  ' parameter];
            
            % obtain settings from a file
            % saving settings
            temp = tempdir;
            if exist(fullfile(temp, 'open_iris_stats_settings.mat'), 'file') == 2
                load(fullfile(temp, 'open_iris_stats_settings.mat'));
                obj.Model.Settings = mibConcatenateStructures(obj.Model.Settings, Settings);    % concatenate Settings structure
                fprintf('Loading settings from %s\n', fullfile(temp, 'open_iris_stats_settings.mat'));
            end
            
            obj.updateWidgets();
			
			% add listner to obj.mibModel and call controller function as a callback
            % option 1: recommended, detects event triggered by mibController.updateGuiWidgets
            obj.listener{1} = addlistener(obj.Model, 'updateGuiWidgets', @(src,evnt) obj.ViewListner_Callback(obj, src, evnt));    % listen changes in number of ROIs
        end
        
        function closeWindow(obj)
            % closing Controller window
            if isvalid(obj.View.gui)
                delete(obj.View.gui);   % delete childController window
            end
            
            % saving settings
            temp = tempdir;
            Settings = obj.Model.Settings;
            save(fullfile(temp, 'open_iris_stats_settings.mat'), 'Settings');
            
            % delete listeners, otherwise they stay after deleting of the
            % controller
            for i=1:numel(obj.listener)
                delete(obj.listener{i});
            end
            
            notify(obj, 'closeEvent');      % notify mibController that this child window is closed
        end
        
        function SelectFilesBtn_Callback(obj)
            % callback to select files
            if isempty(obj.View.ListOfInvoices.Value)
                path = pwd;
            else
                path = fileparts(obj.View.ListOfInvoices.Value);
            end
            [file, path] = uigetfile({'*.xls;*.xlsx','Excel files (*.xls, *.xlsx)'; ...
                '*.*',  'All Files (*.*)'},...
                'Select Excel file', path, 'MultiSelect', 'on');
            if ~iscell(file) && file(1)==0
                return;
            end
            filename = fullfile(path, file);
            if ~iscell(filename)
                filename = {filename};
            end
            obj.View.ListOfInvoices.Items = filename;
            obj.View.ListOfInvoices.Value = filename{1};
            obj.Model.Settings.gui.ListOfInvoices = filename;

            if exist(filename{1}, 'file') == 2
                obj.Model.getColumnNames();
                obj.View.PriceTypeIndexField.Items = sort(obj.Model.VariableNames);
            end

        end
        
        function updateOutputFilename(obj, value)
            % update output filename
            obj.Model.Settings.gui.OutputFilename =  value;
        end
        
        function updateWidgets(obj)
            % function updateWidgets(obj)
            % update widgets of this window
            
            fieldNames = fieldnames(obj.Model.Settings.gui);
            for fieldId = 1:numel(fieldNames)
                if ~isprop(obj.View, fieldNames{fieldId}); continue; end
                switch obj.View.(fieldNames{fieldId}).Type
                    case 'uidropdown'
                        obj.View.(fieldNames{fieldId}).Items = sort({obj.Model.Settings.gui.(fieldNames{fieldId})});
                        obj.View.(fieldNames{fieldId}).Value = obj.Model.Settings.gui.(fieldNames{fieldId});
                    case 'uilistbox'
                        if ~isempty(obj.Model.Settings.gui.(fieldNames{fieldId}))
                            obj.View.(fieldNames{fieldId}).Items = sort(obj.Model.Settings.gui.(fieldNames{fieldId}));
                            obj.View.(fieldNames{fieldId}).Value = obj.Model.Settings.gui.(fieldNames{fieldId}){1};
                        else
                            obj.View.(fieldNames{fieldId}).Items = {};
                        end                        
                    otherwise
                        obj.View.(fieldNames{fieldId}).Value = obj.Model.Settings.gui.(fieldNames{fieldId});
                end
            end
        end
        
        function updateSettings(obj, event)
            if ~isempty(obj.View.PriceTypeIndexField.Items)
                obj.Model.Settings.gui.PriceTypeIndexField = obj.View.PriceTypeIndexField.Value;
            end
            if ~isempty(obj.View.HeaderStartingCell.Value)
                obj.Model.Settings.gui.HeaderStartingCell = obj.View.HeaderStartingCell.Value; 
            end
            if ~isempty(obj.View.DataStartingCell.Value)
                obj.Model.Settings.gui.DataStartingCell = obj.View.DataStartingCell.Value; 
            end
            obj.Model.Settings.gui.OutputFilename = obj.View.OutputFilename.Value;
        end
        
        function StartProcessing(obj)
            % start processing
            obj.View.StartProcessingButton.BackgroundColor = 'r';
            obj.Model.start();
            obj.View.StartProcessingButton.BackgroundColor = 'g';
        end
    end
end