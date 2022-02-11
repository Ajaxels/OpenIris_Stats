function id = findChildId(obj, childName)
% function id = findChildId(childName)
% find id of a child controller
%
% the child controllers of the program are stored in obj.childControllersIds cell
% array. This function look for index that matches with childName string.
% If it is in the list the function returns its index, otherwise it adds it
% to the list as a new element
%
% Parameters:
% childName: name of a child controller
%
% Return values:
% id: index of the requested child controller or empty if it is not open
%

%|
% @b Examples:
% @code id = obj.findChildId('newController');     // find an index of mibImageAdjController @endcode

% Copyright (C) 2019-2020 Ilya Belevich, University of Helsinki (ilya.belevich @ helsinki.fi)
% The MIT License (https://opensource.org/licenses/MIT)
%
% Updates
%

if ismember(childName, obj.childControllersIds) == 0    % not in the list of controllers
    id = [];
else                % already in the list
    id = find(ismember(obj.childControllersIds, childName)==1);
end
end