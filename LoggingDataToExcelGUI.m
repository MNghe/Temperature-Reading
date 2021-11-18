% A MATLAB GUI (Graphical User Interface) that can connect, collect data from
% the 'Arduino Uno' board's thermistor, give a general visual plot of the
% collected data, export the data to an Excel file, delete said Excel
% file, and disconnect from the board.

function varargout = LoggingDataToExcelGUI(varargin)
% LOGGINGDATATOEXCELGUI MATLAB code for LoggingDataToExcelGUI.fig
%      LOGGINGDATATOEXCELGUI, by itself, creates a new LOGGINGDATATOEXCELGUI or raises the existing
%      singleton*.
%
%      H = LOGGINGDATATOEXCELGUI returns the handle to a new LOGGINGDATATOEXCELGUI or the handle to
%      the existing singleton*.
%
%      LOGGINGDATATOEXCELGUI('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in LOGGINGDATATOEXCELGUI.M with the given input arguments.
%
%      LOGGINGDATATOEXCELGUI('Property','Value',...) creates a new LOGGINGDATATOEXCELGUI or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before LoggingDataToExcelGUI_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to LoggingDataToExcelGUI_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help LoggingDataToExcelGUI

% Last Modified by GUIDE v2.5 07-Jun-2021 17:22:07

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @LoggingDataToExcelGUI_OpeningFcn, ...
                   'gui_OutputFcn',  @LoggingDataToExcelGUI_OutputFcn, ...
                   'gui_LayoutFcn',  [] , ...
                   'gui_Callback',   []);
if nargin && ischar(varargin{1})
    gui_State.gui_Callback = str2func(varargin{1});
end

if nargout
    [varargout{1:nargout}] = gui_mainfcn(gui_State, varargin{:});
else
    gui_mainfcn(gui_State, varargin{:});
end
% End initialization code - DO NOT EDIT


% --- Executes just before LoggingDataToExcelGUI is made visible.
function LoggingDataToExcelGUI_OpeningFcn(hObject, eventdata, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to LoggingDataToExcelGUI (see VARARGIN)

% Choose default command line output for LoggingDataToExcelGUI
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes LoggingDataToExcelGUI wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = LoggingDataToExcelGUI_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;


% --- Executes on button press in Connect.
function Connect_Callback(hObject, eventdata, handles)
% hObject    handle to Connect (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
global myComPort;
disp('Connecting... ');
% Connects to the board
myComPort = serial('COM3');
fopen(myComPort)
disp('Connected!');


% --- Executes on button press in exportExcel.
function exportExcel_Callback(hObject, eventdata, handles)
% hObject    handle to exportExcel (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
global Temp;
% Creates the file/directory if not already existing
folder = 'C:\Users\Michael Nghe\Documents\GitHub\MNghe.github.io\Matlab Projs\TempReading';
if ~exist(folder, 'dir')
    mkdir(folder);
end
% File name
baseFileName = 'LoggedData.xlsx';
fullFileName = fullfile(folder, baseFileName);
disp('Creating file... ');
% Create the Excel file / Overwrite the old data/cells
writematrix(Temp', fullFileName);
disp('File created!');


% --- Executes on button press in Disconnect.
function Disconnect_Callback(hObject, eventdata, handles)
% hObject    handle to Disconnect (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
global myComPort;
disp('Disconnecting... ');
% Disconnects from the board
delete(myComPort)
clear myComPort
disp('Disconnected');


% --- Executes on button press in togglebutton.
function togglebutton_Callback(hObject, eventdata, handles)
% hObject    handle to togglebutton (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
global myComPort;
global Temp;

% Checks if the toggle button status is ON
if (get(hObject, 'Value') == 1)
    disp('Starting collection... ');
    % Initialize the data
    Temp = [];
    i = 1;
    % Keeps collecting data until the toggle function is called again
    while get(hObject, 'Value') == 1
        % Collects data
        A = fscanf(myComPort);
        Temp(i) = str2double(string(A));
        i = i + 1;
        if (mod(i, 5) == 0)
            disp('Collecting... ');
        end
        % Gives the program a moment to pause to check if a Callback
        % occurred
        pause (0.01)
    end
% If toggle button is set to OFF
else
    disp('Collection Ended');
    % Creates x-axis for the plot
    x = linspace(0, length(Temp) - 1, length(Temp));
    disp('Graphing... ');
    plot(x, Temp');
end
% Hint: get(hObject,'Value') returns toggle state of togglebutton


% --- Executes on button press in Delete.
function Delete_Callback(hObject, eventdata, handles)
% hObject    handle to Delete (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% Path
folder = 'C:\Users\Michael Nghe\Documents\GitHub\MNghe.github.io\Matlab Projs\TempReading\Excel';
% File name
baseFileName = 'LoggedData.xlsx';
fullFileName = fullfile(folder, baseFileName);
% If file exist, deletes it
if isfile(fullFileName)
    disp('Deleting file... ');
    delete (fullFileName);
    disp('Deleted file');
else
    % Exits the function since the file does not exist
    disp('No Such File Exist');
    return;
end
