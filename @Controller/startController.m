function startController(obj, controllerName, varargin)
% function startController(obj, controllerName, varargin)
% start a child controller using provided name
%
% Parameters:
% controllerName: a string with name of a child controller, for example, 'newController'
% varargin: additional optional controllers or parameters
% varargin{2}: a structure with optional parameters 

%| 
% @b Examples:
% @code obj.startController('newController');     // start a child controller from a callback for handles.mibDisplayBtn press  @endcode
% @code
% Options.opt1 = 1; % [optional] dataset orientation
% Options.opt2 = 1; % [optional] dataset orientation
% obj.startController('newController', [], Options);
% @endcode

% Copyright (C) 2019-2020 Ilya Belevich, University of Helsinki (ilya.belevich @ helsinki.fi)
% The MIT License (https://opensource.org/licenses/MIT)
%
% Updates
%

id = obj.findChildId(controllerName);        % define/find index for this child controller window
if ~isempty(id)
    if numel(varargin) == 2     % run the batch mode when the controller is already opened
        fh = str2func(controllerName);               %  Construct function handle from character vector
        fh(obj.Model, varargin{1:numel(varargin)});
        return;
    else
        try
            figure(obj.childControllers{id}.View.gui);
            obj.childControllers{id}.updateWidgets();   % update widgets of the controller when restarting it
            return; 
        catch err
            obj.childControllersIds(id) = [];
            obj.childControllersIds = obj.childControllersIds(~cellfun('isempty', obj.childControllersIds));
        end
    end
end   % return if controller is already opened

% assign id and populate obj.childControllersIds for a new controller
id = numel(obj.childControllersIds) + 1;    
obj.childControllersIds{id} = controllerName;

fh = str2func(controllerName);               %  Construct function handle from character vector
if nargin > 2 
    obj.childControllers{id} = fh(obj.Model, varargin{1:numel(varargin)});    % initialize child controller with additional parameters
else
    obj.childControllers{id} = fh(obj.Model);    % initialize child controller
end

% add listener to the closeEvent of the child controller
addlistener(obj.childControllers{id}, 'closeEvent', @(src, evnt) mibController.purgeControllers(obj, src, evnt));   % static
%addlistener(obj.childControllers{id}, 'closeEvent', @(src, evnt) obj.purgeControllers(src, evnt)); % dynamic

p = fieldnames(obj.childControllers{id});
if ismember('noGui', p)     % close widgets without GUI
    notify(obj.childControllers{id}, 'closeEvent');
elseif isempty(obj.childControllers{id}.View)   % close widgets with the batch mode
    notify(obj.childControllers{id}, 'closeEvent');
end

end