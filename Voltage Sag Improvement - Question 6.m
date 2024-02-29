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
n = max(max(Y_line{:, 1:2}));                                               % determine the size of the admittance matrix
admittance_matrix = zeros(n);                                               % initialize the admittance matrix

% fill in the off-diagonal elements of the admittance matrix
for i = 1:height(Y_line)
    admittance_matrix(Y_line{i, 1}, Y_line{i, 2}) = -1*Y_line{i, 3};
    admittance_matrix(Y_line{i, 2}, Y_line{i, 1}) = -1*Y_line{i, 3};
end

% fill in the diagonal elements of the admittance matrix
for i = 1:n
    admittance_matrix(i, i) = -1*sum(admittance_matrix(i, :));
end

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

Y = admittance_matrix;

%% Impedance Matrix

% Creating impedance matrix
impedance_matrix = inv(admittance_matrix);                                  % calculate the inverse of the admittance matrix
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

%% Specification for DVR Device

% Assume any value for bus voltage and kVA
kVAbus = 10e3;
Vline = 415;
Vphase = Vline/sqrt(3);

% Calculate phase voltage in unit V for Bus 5 and 14
Vsag5V = Vphase.*Vsag5;
Vsag14V = Vphase.*Vsag14;

% Calculate injection voltage necessary for the Bus 5 and 14
Vinj5 = sqrt(Vphase^2 - Vsag5V.^2);
Vinj14 = sqrt(Vphase^2 - Vsag14V.^2);

% Only take the magnitude of the injected voltage of Bus 5 and 14
Vinj5 = abs(Vinj5);
Vinj14 = abs(Vinj14);

% Calculate the current rating for DVR at Bus 5 and 14
Irating5 = kVAbus/Vphase.*ones(14, 1);
Irating14 = kVAbus/Vphase.*ones(14, 1);

% Calculate the kVA rating for DVR at Bus 5 and 14
Srating5 = 3.*Vinj5.*Irating5/1e3;
Srating14 = 3.*Vinj14.*Irating14/1e3;

% Tabulate all the injection voltage, current and kVA rating for Bus 5 and 14
A = [Vinj5(:,1),Irating5(:,1),Srating5(:,1),Vinj14(:,1),Irating14(:,1),Srating14(:,1)];

%% Question 6

% Display the injected voltage, current and kVA rating of DVR at Bus 5 and 14
disp('                                                             Question 6')
disp(' ')
disp('                                Bus 5                                                          Bus 14')
disp('   ===============================================================================================================================')
disp('    Injected voltage     Current rating (A)     MVA rating (kVA)    Injected voltage     Current rating (A)     MVA rating (kVA)')
disp('   --------------------------------------------------------------------------------------------------------------------------------')
for i = 1:nbus
    fprintf("\t\t%.2f\t\t\t\t\t%.2f\t\t\t\t%.2f\t\t\t\t%.2f\t\t\t\t\t%.2f\t\t\t\t\t%.2f\n",Vinj5(i,1),Irating5(i,1),Srating5(i,1),Vinj14(i,1),Irating14(i,1),Srating14(i,1))
end 
