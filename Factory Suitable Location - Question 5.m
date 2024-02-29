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
Vsag5 = zeros(14,1);
for k = 1:nbus
    Vsag5(k) = buspfv_polar_array(5) - buspfv_polar_array(k)*Z(5,k)/Z(k,k);
end
Vsag5 = Vsag5.';

% Calculate array of voltage sag at bus 14
nbus = length(buspfv_polar_array);
Vsag14 = zeros(14,1);
for k = 1:nbus
    Vsag14(k) = buspfv_polar_array(14) - buspfv_polar_array(k)*Z(14,k)/Z(k,k);
end
Vsag14 = Vsag14.';

% Only take its magnitude to determine the voltage sag magnitude
Vsag5 = abs(Vsag5);
Vsag14 = abs(Vsag14);

%% Expected Fault Occurence for Various Range of Voltages

% Calculate expected fault occurence from linedata
f = linedata{:,"Fault /100 km/year"};
d = linedata{:,"Distance (km)"};
nfault = f.*d/100;
nfault = array2table(nfault, 'VariableNames', {'Expected Fault Occurence'});
linedata = [linedata, nfault(:,1)];

% Expected fault occurence at bus 5
nfault5 = zeros(20, 1);
for i = 1:height(Y_line)
    nfault5(i) = (Vsag5(Y_line{i, 1}) + Vsag5(Y_line{i, 2}))/2;
end
nfault5 = array2table(nfault5, 'VariableNames',{'Expected Fault Occurence at Bus 5'});
nfault5 = [Y_line(:,1:2),nfault5];

% Expected fault occurence at bus 14
nfault14 = zeros(20, 1);
for i = 1:height(Y_line)
    nfault14(i) = (Vsag14(Y_line{i, 1}) + Vsag14(Y_line{i, 2}))/2;
end
nfault14 = array2table(nfault14, 'VariableNames',{'Expected Fault Occurence at Bus 14'});
nfault14 = [Y_line(:,1:2),nfault14];

% Tabulate expected fault occurence of Bus 5 and 14
nfault_bus = [nfault5,nfault14(:,3),nfault(:,1)];

% Calculate voltage sag event at Bus 5 for voltage sag under 0.4 p.u.
fault5no3 = 0;
for i = 1:height(nfault_bus)
    if nfault_bus{i,3} < 0.4
        fault5no3 = fault5no3 + nfault_bus{i,5};
    end
end

% Calculate voltage sag event at Bus 14 for voltage sag under 0.4 p.u.
fault14no3 = 0;
for i = 1:height(nfault_bus)
    if nfault_bus{i,4} < 0.4
        fault14no3 = fault14no3 + nfault_bus{i,5};
    end
end

% Calculate voltage sag event at Bus 5 for various voltage sag percentage of nominal 1.0 p.u.
fault5no = zeros(9,1);
for i = 1:height(nfault_bus)
    if (nfault_bus{i,3} > 0.1) && (nfault_bus{i,3} < 0.2)
        fault5no(1) = fault5no(1) + nfault_bus{i,5};
    elseif (nfault_bus{i,3} > 0.2) && (nfault_bus{i,3} < 0.3)
        fault5no(2) = fault5no(2) + nfault_bus{i,5};
    elseif (nfault_bus{i,3} > 0.3) && (nfault_bus{i,3} < 0.4)
        fault5no(3) = fault5no(3) + nfault_bus{i,5};
    elseif (nfault_bus{i,3} > 0.4) && (nfault_bus{i,3} < 0.5)
        fault5no(5) = fault5no(4) + nfault_bus{i,5};
    elseif (nfault_bus{i,3} > 0.5) && (nfault_bus{i,3} < 0.6)
        fault5no(5) = fault5no(5) + nfault_bus{i,5};
    elseif (nfault_bus{i,3} > 0.6) && (nfault_bus{i,3} < 0.7)
        fault5no(6) = fault5no(6) + nfault_bus{i,5};
    elseif (nfault_bus{i,3} > 0.7) && (nfault_bus{i,3} < 0.8)
        fault5no(7) = fault5no(7) + nfault_bus{i,5};
    elseif (nfault_bus{i,3} > 0.8) && (nfault_bus{i,3} < 0.9)
        fault5no(8) = fault5no(8) + nfault_bus{i,5};
    elseif (nfault_bus{i,3} > 0.9) && (nfault_bus{i,3} < 1.0)
        fault5no(9) = fault5no(9) + nfault_bus{i,5};
    end
end

% Calculate voltage sag event at Bus 14 for various voltage sag percentage of nominal 1.0 p.u.
fault14no = zeros(9,1);
for i = 1:height(nfault_bus)
    if (nfault_bus{i,4} > 0.1) && (nfault_bus{i,4} < 0.2)
        fault14no(1) = fault14no(1) + nfault_bus{i,5};
    elseif (nfault_bus{i,4} > 0.2) && (nfault_bus{i,4} < 0.3)
        fault14no(2) = fault14no(2) + nfault_bus{i,5};
    elseif (nfault_bus{i,4} > 0.3) && (nfault_bus{i,4} < 0.4)
        fault14no(3) = fault14no(3) + nfault_bus{i,5};
    elseif (nfault_bus{i,4} > 0.4) && (nfault_bus{i,4} < 0.5)
        fault14no(4) = fault14no(4) + nfault_bus{i,5};
    elseif (nfault_bus{i,4} > 0.5) && (nfault_bus{i,4} < 0.6)
        fault14no(5) = fault14no(5) + nfault_bus{i,5};
    elseif (nfault_bus{i,4} > 0.6) && (nfault_bus{i,4} < 0.7)
        fault14no(6) = fault14no(6) + nfault_bus{i,5};
    elseif (nfault_bus{i,4} > 0.7) && (nfault_bus{i,4} < 0.8)
        fault14no(7) = fault14no(7) + nfault_bus{i,5};
    elseif (nfault_bus{i,4} > 0.8) && (nfault_bus{i,4} < 0.9)
        fault14no(8) = fault14no(8) + nfault_bus{i,5};
    elseif (nfault_bus{i,4} > 0.9) && (nfault_bus{i,4} < 1.0)
        fault14no(9) = fault14no(9) + nfault_bus{i,5};
    end
end



%% Question 5

% Display voltage sag at bus 5 and bus 14
disp('Question 5')
disp(' ')

% Determining which bus is the most suitable to be connected to factory
if (fault5no3 < fault14no3) && (fault5no(8) > fault14no(8))
    fprintf('Bus 5 is the most suitable place to place a factory manufacturing electronic component.\n')
    fprintf('The number of voltage sag event occurring under 0.4 p.u. of the nominal 1.0 p.u. at Bus 5 is %.2f, which is lower than Bus 14, which is %.2f.\n',fault5no3,fault14no3)
    fprintf('the number of voltage sags event at Bus 5 when three-phase fault occurs at every buses are more frequently to have the voltage sag magnitude between \n0.8 to 0.9 p.u., which is closer to nominal 1.0 p.u. if compared to Bus 14, where the voltage sag magnitude for every voltage sags event are inconsistently \ndistributed between 0.1 p.u. to 0.8 p.u. and less likely to occurs with voltage sag magnitude between 0.8 and 0.9\n')
elseif (fault5no3 > fault14no3) && (fault5no(8) < fault14no(8))
    fprintf('Bus 14 is the most suitable place to place a factory manufacturing electronic component.\n')
    fprintf('The number of voltage sag event occurring under 0.4 p.u. of the nominal 1.0 p.u. at Bus 14 is %.2f, which is lower than Bus 5, which is %.2f.\n',fault14no3,fault5no3)
    fprintf('the number of voltage sags event at Bus 14 when three-phase fault occurs at every buses are more frequently to have the voltage sag magnitude between \n0.8 to 0.9 p.u., which is closer to nominal 1.0 p.u. if compared to Bus 5, where the voltage sag magnitude for every voltage sags event are inconsistently \ndistributed between 0.1 p.u. to 0.8 p.u. and less likely to occurs with voltage sag magnitude between 0.8 and 0.9\n')
end

