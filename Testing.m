% Reading the full file path to the Excel file
filename = '/Users/maryclaireokoro/Downloads/HW1_2024_TowerData-2.xlsx';

% Read the first sheet of the Excel file into a table
dataTable = readtable(filename, 'Sheet', 1);

% Display the first few rows of the table
disp(dataTable);

% Extracting wind speed, wind direction and time from the table
wind_speed = dataTable.WindSpeed_50_2M_;
wind_direction = dataTable.Direction_Deg_;
time = dataTable.ActualDatesAndTimes;


% Obtaining the mean wind speed of the tower data
mean_wind_speed = mean(wind_speed);

disp(['Mean Wind Speed:', num2str(mean_wind_speed), 'm/s']);

% Calculating the standard deviation of the wind speed
std_wind_speed = std(wind_speed);

disp(['Standard Deviation of Wind Speed:', num2str(std_wind_speed), 'm/s']);

% Question 2 - Plot wind speed vs time
figure;
plot(time, wind_speed);
xlabel('Time');
ylabel('Wind Speed (m/s)');
title('Wind Speed versus Time');

% Question 3 - Plot wind direction vs time
figure;
plot(time, wind_direction);
xlabel('Time');
ylabel('Wind Direction (Degrees)');
title('Wind Direction versus Time');

% Question 4 - Creating a Wind-rose
% Converting wind direction to radians for polarhistogram
wind_direction_rad = deg2rad(wind_direction);

% Creating a polar histogram for wind direction
polarhistogram(wind_direction_rad, 16, 'FaceColor', 'blue'); % 16 sectors of 22.5

% Adjusting plot to North at the top
set(gca, 'ThetaZeroLocation', 'top', 'ThetaDir', 'clockwise');
title('Wind-rose - Wind Direction and Speed');


% Question 5 - Calculating a pdf wind speed curve
figure;
histogram(wind_speed, 'Normalization', 'pdf');
xlabel('Wind Speed (m/s)');
ylabel('Probability Density');
title('PDF of Wind Speed');

% Question 6 - 
% Wind speed and power output(kw) for the Vestas V52/850 turbine
wind_speed_curve = [0,3,3.5,4,4.5,5,5.5,6,6.5,7,7.5,8,8.5,9,9.5,10,10.5,11,11.5,12,12.5,13,13.5,14,25];
power_output_curve = [0,5,15,30,49,67,97,127,162,197,244,290,350,410,475,539,600,660,710,751,784,816,840,850,850];

% Interpolating power curve 
power_curve_interpolation = @(v) interp1(wind_speed_curve, power_output_curve, v, 'linear', 0);

% Extracting PDF values 
[wind_speed_counts, wind_speed_edges] = histcounts(wind_speed, 'Normalization','pdf');

% Computing bin centers for wind speed range
wind_speed_range = (wind_speed_edges(1:end-1) + wind_speed_edges(2:end)) /2;

% Getting the PDF values
pdf_values = wind_speed_counts;

% Calculating expected power for each wind speed
expected_power = power_curve_interpolation(wind_speed_range) .* pdf_values;

% Plotting expected power versus wind speed
figure;
plot(wind_speed_range, expected_power);
xlabel('Wind Speed (m/s)');
ylabel('Expected Power');
title('Expected Power versus Wind Speed');

% Integrating expected power over wind speed range to get average power
average_power_kW = trapz(wind_speed_range, expected_power);

% Multiplying by the number of hours in a year to get total energy produced (kWh)
hours_per_year = 8760;
total_energy_kWh = average_power_kW * hours_per_year;

disp(['Estimated Total Energy Production (kWh/year): ', num2str(total_energy_kWh)]);
