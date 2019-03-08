%% Calcium Analysis
% *Clear the workspace*

clear all
clc
%% 
% Import excel spreadsheet from NIS Elements with data, in the following 
% format: 
% 
% Frame    Cell 1    Cell 2    Etc
% 
% 1            Value    Value    Value
% 
% 2            Value    Value    Value

File = uigetfile('*.xlsx');
[Pre_data, Pre_headers] = xlsread(File);
Pre_headers = Pre_headers(2,:);
Data = Pre_data(:,3:end-1);
Data_headers = Pre_headers(1, 3:end-1);
Frame_column = Pre_data(:,1);
Frame_header = Pre_headers(1,1);
Time_column = Pre_data(:,2);
Time_header = Pre_headers(1,2);
[NumberofRows, NumberofColumns] = size(Data);
Max_value = max(max(Data));
Pre_time = datestr(Time_column, 'MM:SS:FFF');
Time = cellstr(Pre_time);
%%
%Cell_type = questdlg('What is the cell type?', 'Cell type', 'Neuron', 'Astrocyte', 'Cancel', 'Cancel');
%Cell_type = cellstr(Cell_type);%

Treatment = questdlg('What is the treatment?', 'Treatment', 'Baseline', 'KA', 'Cancel', 'Cancel');
Treatment = cellstr(Treatment);

Specifications = inputdlg({'What is the Short ID?', 'What is the Slice?', 'What is the genotype?', 'What is the age?', 'What is the full Animal ID?'}, 'Specifications');

fps = inputdlg('What is the fps rate?');
fps = str2double(fps);
fps = round(fps);
%% 
% *Plot Raw Data*

Figure1 = figure;
plot(Frame_column, Data)
axis([0 NumberofRows 0 Max_value+10])
title('Raw signal')
legend(Data_headers, 'Location', 'eastoutside')
ylabel('Intensity')
xlabel('Frame')
%% 
% *Plot Individual Signal*

%Find out number dimension of figure for subplots
% according to how many columns there are, by
% finding the square root of the # of columns,
% approximated by ceil function

Dimensionofsubplots = ceil(sqrt(NumberofColumns));
Figure2 = figure;
for i = 1:NumberofColumns
    subplot1 = subplot(Dimensionofsubplots, Dimensionofsubplots, i);
    plot(Data(:,i));
    
    %plot specifications
    Max_value_subplot = max(max(Data(:,i)));
    axis([0 NumberofRows 0 Max_value_subplot+30]);
    title(Data_headers(i));
    xlabel('Frame');
    ylabel('Intensity', 'FontSize', 8)
end
%% 
% *Calculate the overall mean*

Figure3 = figure;
OverallMeans = mean(Data);
bar(OverallMeans);

%plot specifications
title('Average of Gray Value');
xlabel('Cell #');
ylabel('Intensity');
%% 
% *Calculate the average of first second *

Mean = mean(Data(1:8, :))
%% *Calculate dF/F (normalized data)*
%%
DeltaFoverF = (Data - Mean)./Mean;

% Finds hidden mistakes in DF/F
for i = 1:NumberofColumns
    for j = 1:NumberofRows
        if DeltaFoverF(j,i)==9 | isnan(DeltaFoverF(j,i))
           error('Something is wrong')          
        end
    end
end

%%
%plot with both positive and negative values
plot(DeltaFoverF)
%plot specifications
axis([0 NumberofRows ylim])
title('\DeltaF/F');
xlabel('Frame');
ylabel('Intensity');

%Plot DeltaFoverF of each cell
Figure2 = figure;
for i = 1:NumberofColumns
    subplot1 = subplot(Dimensionofsubplots, Dimensionofsubplots, i);
    plot(DeltaFoverF(:,i));
    
    %plot specifications
    Max_value_subplot = max(max(DeltaFoverF(:,i)));
    title(Data_headers(i));
    axis([0 NumberofRows min(min(DeltaFoverF)) max(max(DeltaFoverF))])
    xlabel('Frame');
    ylabel('Intensity', 'FontSize', 8)
end

%%
%Delta F with positive values only, negative values transformed in 0
Figure3 = figure;
DeltaFpositivesonly = DeltaFoverF;
DeltaFpositivesonly(DeltaFpositivesonly<0) = 0;
plot(DeltaFpositivesonly)

%plot specifications
axis([0 NumberofRows ylim])
title('\DeltaF/F positives only');
xlabel('Frame');
ylabel('Intensity');


%% 
% *Calculate the standard deviation of previous 10 frames*

StdDev = movstd(DeltaFoverF, [9 0]);
StandardDeviationPrevious10 = StdDev(10:end, :);
StandardDeviation = 2.5 .* StandardDeviationPrevious10;               % Creates array with values of 2.5*STD of 10 previous frames
%% Find all peaks and valleys
%%
DeltaFoverF2 = DeltaFoverF(10:end, :);
[NumberofRows3, ~] = size(DeltaFoverF2);
Frame_column2 = 20:(NumberofRows3+19);
Frame_column3 = 1:NumberofRows3;

%Finds all peaks and all valleys in DeltaFoverF2
for i = 1:NumberofColumns
    [allpeaks{i}, allpeaklocation{i}] = findpeaks(+DeltaFoverF2(:,i));
    [allvalleys{i}, allvalleylocation{i}] = findpeaks(-DeltaFoverF2(:,i));
    allvalleysneg{1,i}= -allvalleys{1,i};
end

%Example in graph of all peaks and valleys on DeltaFoverF2 column 1
plot(Frame_column3, DeltaFoverF2(:,2))
hold on
plot(allpeaklocation{1,2}, allpeaks{1,2}, '*m');
plot(allvalleylocation{1,2}, allvalleysneg{1,2}, 'sq');
hold off

% All graphs of all peaks and valleys
Figure4 = figure;
for i = 1:NumberofColumns
    subplot2 = subplot(Dimensionofsubplots, Dimensionofsubplots, i);
    plot(Frame_column3, DeltaFoverF2(:,i))
    hold on
    plot(allpeaklocation{1,i}, allpeaks{1,i}, '*m');
    plot(allvalleylocation{1,i}, allvalleysneg{1,i}, 'sq');
    hold off
    
    %plot specifications
    Max_value_subplot = max(max(DeltaFoverF2(:,i)));
    axis([0 NumberofRows3 -2 Max_value_subplot+1]);
    title(Data_headers(i));
    xlabel('Frame');
    ylabel('Intensity', 'FontSize', 8)  
end
%% 
% *Find peaks bigger than the threshold (Valid peaks)*

% Creates matrix with values of DealtaF/F bigger than 2.5x StDev of 10 previous frames, otherwise, Zeros
for i = 1:NumberofColumns
    for j = 1:NumberofRows3
        if DeltaFoverF2(j,i) >= StandardDeviation(j,i);
            Peakarray(j,i) = DeltaFoverF2(j,i);
        else 
            Peakarray(j,i) = 0;
        end
    end
end

% Creates array with locations where DF/F is bigger than 2.5x StDev of 10 prev frames.
for i = 1:NumberofColumns
    Valid_peak_locations{i} = find(Peakarray(:,i)>0);
end

% Creates array with values stored in locations above; the real peaks.
for i = 1:NumberofColumns
    Valid_peaks{i} = nonzeros(Peakarray(:,i));
end

% Plot valid peaks
Figure5 = figure;
for i = 1:NumberofColumns
    subplot3 = subplot(Dimensionofsubplots, Dimensionofsubplots, i);
    plot(Frame_column3, DeltaFoverF2(:,i))
    hold on
    plot(Valid_peak_locations{1,i}, Valid_peaks{1,i}, '*m');
    hold off
    
    %plot specifications
    Max_value_subplot = max(max(DeltaFoverF2(:,i)));
    axis([0 NumberofRows3 -2 Max_value_subplot+1]);
    title(Data_headers(i));
    xlabel('Frame');
    ylabel('Intensity', 'FontSize', 8)  
end
%% 
% *Finds valleys before and after valid peaks*

% Finds valleys before and after all the peaks
% Calculates width, risetime and decaytime of all peaks

for i = 1:NumberofColumns
    for j = 1:(length(allpeaklocation{1,i})-1)        %%% -1 to make allpeak have same n of elements as allvalley
        valley_before{j,i} = allvalleylocation{1,i}(find(allvalleylocation{1,i}<allpeaklocation{1,i}(j,1), 1, 'last'));
        valley_after{j,i} = allvalleylocation{1,i}(find(allvalleylocation{1,i}>allpeaklocation{1,i}(j,1), 1, 'first'));
        width{j,i} = valley_after{j,i} - valley_before{j,i};
        risetime{j,i} = allpeaklocation{1,i}(j,1) - valley_before{j,i};
        decaytime{j,i} = valley_after{j,i} - allpeaklocation{1,i}(j,1);
    end
end

% Finds location of valid peaks info in other matrixes

for i = 1:NumberofColumns
    for j = 1:(length(Valid_peak_locations{1,i}))
        Index_of_Validpeaks{j,i} = find(allpeaklocation{1,i}==Valid_peak_locations{1,i}(j,1));
    end
end

% Transforms cells within cells into arrays

for i = 1:NumberofColumns
    for j = 1:(length(Index_of_Validpeaks))
        pValid_valley_before_location{j,i} = valley_before(Index_of_Validpeaks{j,i});
        pValid_valley_after_location{j,i} = valley_after(Index_of_Validpeaks{j,i});
        pValid_width{j,i} = width(Index_of_Validpeaks{j,i});
        pValid_risetime{j,i} = risetime(Index_of_Validpeaks{j,i});
        pValid_decaytime{j,i} = decaytime(Index_of_Validpeaks{j,i});
    end
end

% It fixes cell within a cell 

for i = 1:NumberofColumns
    Valid_valley_before_location{i} = vertcat(pValid_valley_before_location{:,i});
    Valid_valley_after_location{i} = vertcat(pValid_valley_after_location{:,i});
    Valid_width{i} = vertcat(pValid_width{:,i});
    Valid_risetime{i} = vertcat(pValid_risetime{:,i});
    Valid_decaytime{i} = vertcat(pValid_decaytime{:,i});
end

% Gets the index right before the wanted index, to find the "valley before" later
for i = 1:NumberofColumns
    for j = 1:(length(Index_of_Validpeaks))
        Index_minus_one{j,i} = Index_of_Validpeaks{j,i} - 1;
    end
end

%%
for i = 1:NumberofColumns
    for j = 1:(length(Index_of_Validpeaks))
        if Index_minus_one{j,i} ~= 0
            Valid_valley_after{j,i} = allvalleys{1,i}(Index_of_Validpeaks{j,i});
            Valid_valley_before{j,i} = allvalleys{1,i}(Index_minus_one{j,i});
        else
            Valid_valley_after{j,i} = allvalleys{1,i}(Index_of_Validpeaks{j,i});
            Valid_valley_before{j,i} = 0;
        end
        Valid_valley_before{j,i} = -Valid_valley_before{j,i};
        Valid_valley_after{j,i} = -Valid_valley_after{j,i};
    end
end

for i = 1:NumberofColumns
Valid_valley_before_location_mat{i} = cell2mat(Valid_valley_before_location{1,i});
Valid_valley_after_location_mat{i} = cell2mat(Valid_valley_after_location{1,i});
Valid_valley_before_mat{i} = vertcat(Valid_valley_before{:,i});
Valid_valley_after_mat{i} = vertcat(Valid_valley_after{:,i});
end

%%
%Example in graph of original signal and clean signal
    % Original = DeltaFpositives only
    % Clean = Peakarray values
plot(Frame_column3, DeltaFoverF2(:,5))
hold on
plot(Valid_peak_locations{1,5}, Valid_peaks{1,5}, 'sq')
plot(Valid_valley_before_location_mat{1,5}, Valid_valley_before_mat{1,5}, '*')
plot(Valid_valley_after_location_mat{1,5}, Valid_valley_after_mat{1,5}, 'x')

%% 
% *Display the clean signal (only with 'valid' spikes) of each cell*

%Plot all cells' clean signal
[NumberofRows6, ~] = size(Peakarray);
Dimensionofsubplots = ceil(sqrt(NumberofColumns6));
Figure6 = figure;
for i = 1:NumberofColumns
    subplot2 = subplot(Dimensionofsubplots, Dimensionofsubplots, i);
    plot(Peakarray(:,i));
    
    %plot specifications
    Max_value_subplot = max(max(Peakarray(:,i)));
    axis([0 NumberofRows6 0 Max_value_subplot]);
    title(Data_headers(i));
    xlabel('Frame');
    ylabel('Intensity', 'FontSize', 8)
end
%% 
% *Average of Peaks amplitudes*

%Average of peaks amplitudes
for i = 1:NumberofColumns
    AvgofPeaksAmplitudes{i} = mean(Peakamplitude{1,i});
end

%Transform cell array with averages into matlab array
AvgofPeaksAmplitudes = cell2mat(AvgofPeaksAmplitudes);
%% Frequency of spikes

Duration = duration(00,05,00);

for i = 1:NumberofColumns
    Frequency{i} = length(Peaklocation{1,i})/5;
end

Frequency = cell2mat(Frequency);
Frequency = Frequency';
%% 
% *Write results to Excel file*

%Write results into a table
Cell_ID = Data_headers';
Genotype = repmat(Specifications(3,1), [length(Cell_ID), 1]);
Age = repmat(Specifications(4,1), [length(Cell_ID), 1]);
Short_ID = repmat(Specifications(1,1), [length(Cell_ID), 1]);
Animal_ID = repmat(Specifications(5,1), [length(Cell_ID), 1]);
Slice = repmat(Specifications(2,1), [length(Cell_ID), 1]);
Cell_type = repmat(Cell_type, [length(Cell_ID), 1]);
Treatment = repmat(Treatment, [length(Cell_ID), 1]);
Average_of_Spikes_Amplitudes = AvgofPeaksAmplitudes';
Frequency_of_Spikes = Frequency';

Table = table(Genotype, Age, Short_ID, Slice, Animal_ID, Cell_ID, Cell_type, Treatment, Average_of_Spikes_Amplitudes, Frequency)

%Write table to Excel file
filename = uiputfile('*.xlsx', 'Save to Excel');
writetable(Table, filename)
