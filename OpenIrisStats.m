function OpenIrisStats()
% @mainpage OpenIrisStats
% @section intro Introduction
% @b OpenIrisStats is a Matlab based program for generation of usage statistics from the invoices
% produced with the OpenIris reservation system (https://openiris.io).
% @section features Generation of statistics for
% - users
% - usage hours
% @section description Description
% Statistcis

% Copyright (C) 2022-2023 Ilya Belevich, University of Helsinki (ilya.belevich @ helsinki.fi)
% The MIT License (https://opensource.org/licenses/MIT)

% turn off warnings
warning('off', 'MATLAB:ui:javaframe:PropertyToBeRemoved'); 

if ~isdeployed
    func_name='OpenIrisStats.m';
    func_dir=which(func_name);
    func_dir=fileparts(func_dir);
    addpath(func_dir);
    addpath(fullfile(func_dir, 'Tools'));
    % addpath(fullfile(func_dir, 'Classes'));
end
version = 'ver. 2025.01 (23.01.2025)';
if isdeployed; version = [version ' Academic version']; end
model = Model();     % initialize the model
controller = Controller(model, version);  % initialize controller
