% NAME       : MUHAMMAD ZIKRIL HAKIM BIN ZULKIFLY
% MATRIX NO. : 17187156/2

clc; close all; clear;
%% Data Extraction

% Extract data from given value in the question
filename = 'Data.xlsx';
sheet1 = 'Line Data';
sheet2 = 'Generator Data';
sheet3 = 'Buses Pre-Fault Voltage';
header1 = 'A1:G1';
header2 = 'A1:D1';
header3 = 'A1:C1';
range1 = 'A2:G21';
range2 = 'A2:D6';
range3 = 'A2:C15';

% Creating a table in MATLAB
[~, headers1] = xlsread(filename, sheet1, header1);
[~, headers2] = xlsread(filename, sheet2, header2);
[~, headers3] = xlsread(filename, sheet3, header3);
data1 = xlsread(filename, sheet1, range1);
data2 = xlsread(filename, sheet2, range2);
data3 = xlsread(filename, sheet3, range3);
linedata = array2table(data1, 'VariableNames', headers1);
gendata = array2table(data2, 'VariableNames', headers2);
buspfv = array2table(data3, 'VariableNames', headers3);

%% Admittance Matrix

% Creating Y admittance matrix
R = linedata{:, 3};                                                         % extract the 3rd column of linedata as an array
X = linedata{:, 4};                                                         % extract the 4th column of linedata as an array
B = linedata{:, 5};                                                         % extract the 5th column of linedata as an array

% Admittance calculation
Y_line = 1./(R + 1j*X) + 1j*B;                                              % calculate Y_line using the given formula
Y_line = array2table(Y_line, 'VariableNames', {'Y'});                       % convert Y_line to a table
Y_line = [linedata(:, 1:2), Y_line];                                        % insert the first and second columns of line data into Y

% Initialize the admittance matrix
n = max(max(Y_line{:, 1:2})); % determine the size of the admittance matrix
admittance_matrix = zeros(n); % initialize the admittance matrix

% fill in the off-diagonal elements of the admittance matrix
for i = 1:height(Y_line)
    admittance_matrix(Y_line{i, 1}, Y_line{i, 2}) = -1*Y_line{i, 3};
    admittance_matrix(Y_line{i, 2}, Y_line{i, 1}) = -1*Y_line{i, 3};
end

% fill in the diagonal elements of the admittance matrix
for i = 1:n
    admittance_matrix(i, i) = -1*sum(admittance_matrix(i, :));
end

% Deleting the information about Generator 3 and Generator 5
gendata(3, :) = [];
gendata(4, :) = [];

% Extract data from table gendata
gen_bus = gendata{:,2};
gen_reactance = gendata{:,3};
gen_mva = gendata{:,4};

% Calculating new generator reactance
gen_reactance = gen_reactance.*(100./gen_mva);
newgenX = array2table(gen_reactance, 'VariableNames', {'New X"'});
gendata = [gendata, newgenX(:,1)];

% Convert reactance to admittance
Ygen = 1./(1j*gen_reactance);

% Add generator admittance to admittance matrix
for i = 1:length(gen_bus)
    admittance_matrix(gen_bus(i), gen_bus(i)) = admittance_matrix(gen_bus(i), gen_bus(i)) + Ygen(i);
end

% Substitute admittance matrix to Y
Y = admittance_matrix;

%% Impedance Matrix

% Creating impedance matrix
impedance_matrix = inv(Y);                                                  % calculate the pseudo-inverse of the admittance matrix
Z = impedance_matrix;

%% Bus Pre-Fault Voltage

% Extract data from table buspfv
voltage_magnitude = buspfv{:,2};
voltage_angle = buspfv{:,3};

% Convert angle from degrees to radians
voltage_angle = deg2rad(voltage_angle);

% Combine magnitude and angle into polar form
voltage_polar = voltage_magnitude .* exp(1i * voltage_angle);
buspfv_polar = array2table(voltage_polar, 'VariableNames', {'Pre-fault voltage (p.u.)'});
buspfv_polar = [buspfv(:, 1), buspfv_polar];
buspfv_polar_array = table2array(buspfv_polar(:,2));

%% Voltage Sag Calculation

% Calculate array of voltage sag at bus 5
nbus = length(buspfv_polar_array);
for k = 1:nbus
    Vsag5(k) = buspfv_polar_array(5) - buspfv_polar_array(k)*Z(5,k)/Z(k,k);
end
Vsag5 = Vsag5.';

% Calculate array of voltage sag at bus 14
nbus = length(buspfv_polar_array);
for k = 1:nbus
    Vsag14(k) = buspfv_polar_array(14) - buspfv_polar_array(k)*Z(14,k)/Z(k,k);
end
Vsag14 = Vsag14.';

% Only take its magnitude to determine the voltage sag magnitude
Vsag5 = abs(Vsag5);
Vsag14 = abs(Vsag14);

%% Question 2(b)

% Display voltage sag at bus 5 and bus 14
disp('     Question 2(b)')
disp(' ')
disp('    Bus 5     Bus 14')
disp('   ==================')
disp([Vsag5               Vsag14])