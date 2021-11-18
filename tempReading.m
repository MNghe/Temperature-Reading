% This script will connect to a physically connected microcontroller, an
% Elegoo Uno R3 in my case, that is identical to an Arduino Uno board. The
% board is connected to a thermistor to collect temperature data in the
% surroundings. Once the data is collected, based on n number of samples,
% it will then plot the data and create a Microsoft Excel file to store the
% data if not already created.

% Clears Command Window history & Workspace
clear;
clc;

% Use 'serialportlist' in Command Window to find available ports
% In this case, it is 'COM3'
myComPort = serial('COM3');

% Disconnects the device in the case that it is still connected
delete(myComPort)
clear myComPort

% Now connect the device
myComPort = serial('COM3');
fopen(myComPort)

% n represents the number of samples
n = 100;

% The x-axis for the plot
x = linspace(0, n-1, n);

% Gets the data from the 'Arduino Uno' board and stores it into an array
for i = 1:n
    Temp(i) = str2double(string(fscanf(myComPort)));
end

% Now disconnect the device
delete(myComPort)
clear myComPort

% Plots the graph
plot(x,Temp);

% If this folder/directory does not exist, creates the directory
folder = 'C:\Users\Michael Nghe\Documents\MATLAB\Matlab Projs\Excel';
if ~exist(folder, 'dir')
    mkdir(folder);
end
% Name of the Microsoft Excel file
baseFileName = 'LoggedData.xlsx';

% Creates the Microsoft Excel file / overwrites the old data/cells
fullFileName = fullfile(folder, baseFileName);
writematrix(Temp', fullFileName);