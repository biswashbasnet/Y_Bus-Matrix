%Biswash_Basnet( This script provides different options for entering bus data)
clc;
clear all;
warning('off', 'all');
disp('Choose an option:');
disp('1. Use existing predefined data in code');
disp('2. Load from an Excel/CSV file in the format given within code');
disp('3. Enter data manually (Bus name and Bus data)');
choice = input('Enter your choice (1/2/3): ', 's');
if strcmp(choice, '1')  % Use predefined existing data
    predefined_data = { ...
        1, 'Birch', 'Elm', 0.042, 0.168;
        2, 'Birch', 'Pine', 0.031, 0.126;
        3, 'Elm', 'Maple', 0.031, 0.126;
        4, 'Maple', 'Oak', 0.084, 0.336;
        5, 'Maple', 'Pine', 0.053, 0.21;
        6, 'Oak', 'Pine', 0.063, 0.252; 
    };
    data = cell2table(predefined_data, 'VariableNames', {'LineNo', 'StartBus', 'EndBus', 'R', 'X'});  % changing to table similar to Excel format
elseif strcmp(choice, '2')  % Load from a CSV/XLS file
    filename = input('Enter the name of the file (with extension): ', 's');
    try
        [~, ~, ext] = fileparts(filename);
        if strcmp(ext, '.csv')
            opts = detectImportOptions(filename, 'NumHeaderLines', 1); % simpler import for CSV and Skip header row
            data = readtable(filename, opts); % Read CSV file correctly
        elseif strcmp(ext, '.xls') || strcmp(ext, '.xlsx')
            data = readtable(filename, 'Sheet', 1); %Read first sheet of the Excel file
        else
            error('Unsupported file format. Use .csv or .xls/.xlsx.');
        end
    catch ME
        disp(['Error: ', ME.message]);
        return;
    end
elseif strcmp(choice, '3')  % Manual Entry
    disp('Enter the names of buses in the transmission lines:');
    buses = {};
    i = 1;
    while true
        buses{i} = input(['Enter the name of Bus ', num2str(i), ': '], 's');
        addAnother = lower(input('Is there another bus? (yes/no): ', 's'));
        if strcmp(addAnother, 'no')
            break;
        end
        i = i + 1;
    end
    Nbus = length(buses);
    bus_map = containers.Map(buses, 1:Nbus);
    disp(['Total Buses: ', num2str(Nbus)]);
    linedata = [];
    Nline = 0;
    for i = 1:Nbus-1
        for k = i+1:Nbus
            connection = lower(input(['Connection between ', buses{i}, ' and ', buses{k}, ' (yes/no)?: '], 's'));
            while ~strcmp(connection, 'yes') && ~strcmp(connection, 'no')
                disp('Invalid input. Enter "yes" or "no".');
                connection = lower(input(['Connection between ', buses{i}, ' and ', buses{k}, ' (yes/no)?: '], 's'));
            end
            if strcmp(connection, 'yes')
                Nline = Nline + 1;
                R = input(['Enter resistance (p.u.) between ', buses{i}, ' and ', buses{k}, ': ']);
                X = input(['Enter reactance (p.u.) between ', buses{i}, ' and ', buses{k}, ': ']);
                linedata(Nline, :) = [Nline, i, k, R, X];
            end
        end
    end
else
    disp('Invalid choice.');
    return;
end
if exist('data', 'var') % Process Data from Table or Manual Input
    start_buses = data{:, 2}; 
    end_buses = data{:, 3};   
    resistances = data{:, 4};  
    reactances = data{:, 5};   
    unique_buses = unique([start_buses; end_buses]);
    Nbus = length(unique_buses);
    bus_map = containers.Map(unique_buses, 1:Nbus); 
    Nline = height(data);
    linedata = zeros(Nline, 5);
    for i = 1:Nline
        linedata(i, :) = [data{i, 1}, bus_map(start_buses{i}), bus_map(end_buses{i}), resistances(i), reactances(i)];
    end
end
disp(['Total Buses: ', num2str(Nbus)]);% **Display Final Data**
disp(['Total Lines: ', num2str(Nline)]);
disp('-------------------------------------------------------------------------------------------------')
disp(' SN. No  | Start Bus (No. Name) | End Bus (No. Name) | Resistance (p.u.) | Reactance (p.u.) ')
disp('-------------------------------------------------------------------------------------------------')
bus_names = keys(bus_map);% Reverse map to get bus names from numbers
bus_numbers = values(bus_map);
bus_reverse_map = containers.Map(bus_numbers, bus_names);
for row = 1:Nline
    start_bus_no = linedata(row, 2);
    end_bus_no = linedata(row, 3);
    start_bus_name = bus_reverse_map(start_bus_no);
    end_bus_name = bus_reverse_map(end_bus_no);
    fprintf('   %-6d |   %-2d %-7s    |   %-2d %-7s   |       %-12.6f |       %-12.6f \n', ...
            linedata(row, 1), start_bus_no, start_bus_name, end_bus_no, end_bus_name, linedata(row, 4), linedata(row, 5));
end
disp('-------------------------------------------------------------------------------------------------')
Ybus = zeros(Nbus, Nbus); % Ybus calculations
for i = 1:Nline
    p = linedata(i, 2);
    q = linedata(i, 3);
    Z = linedata(i, 4) + 1j * linedata(i, 5);
    if Z ~= 0  % Avoid division by zero
        yline = 1/Z;
        Ybus(p, p) = Ybus(p, p) + yline;
        Ybus(q, q) = Ybus(q, q) + yline;
        Ybus(p, q) = Ybus(p, q) - yline;
        Ybus(q, p) = Ybus(q, p) - yline;
    end
end
disp('Initial Ybus matrix:');
disp(Ybus);
input('Press Enter to exit...', 's');
