function [Mat] = ExtractMatrix(Freq, InputFileAdd)
%this function gets an excel file and extracts the matrix of the requested
%frequency
%Mat - the matrux that will be returned
%Freq - requested frequency
%InputFileAdd - input file address - name & path
%skip - if we find the right frequency we return the full matrix and 
%       skip = 0, otherwise we return an empty matrix and skip = 1

%%
skip = 0;
%% Ensuring all inputs are valid
if isempty(Freq)
    fprintf('Error: Please enter frequency for figures\n')
    skip = 1; %if skip changes to one the whole function will break
end
if isempty(InputFileAdd)
    fprintf('Error: Please enter input file address\n')
    skip = 1; %if skip changes to one the whole function will break
end

if (skip == 0)
    %get the names of all the sheets
    [name, sheet] = xlsfinfo(InputFileAdd);
    %go through all the sheets looking for the string 'GHz'
    for i=1:length(sheet)
        findGHz(i) = strfind(sheet(i),'GHz');
    end
    %if you didn't find any 'GHz':
    if isempty(findGHz)
        fprintf('Congratulations! You fucked shit up again!\nWhy dont you try entering an appropriate Excel file\n');
    else
        %if you did find - extract the frequencies from the file
        j=1;
        for i=1:length(sheet(:))
            if isempty(findGHz{i})
                %do nothing
            else
                %change format of sheet name
                sheetName = cell2mat(sheet(i));
                %build a matrix of all frequencies
                f(j) = str2num(sheetName(1:findGHz{i}-2));
                if(f(j) == Freq)
                    data = xlsread(InputFileAdd, i);
                    %find all the non-NaN data and it's indices
                     [row,col] = find( ~isnan(data((2:length(data(:,1))),(2:length(data(1,:))))));
                     %marking the 2 indices that limit our non-NaN data
                     el = data(2:length(data(:,1)), 1);
                     az = data(1, 2:length(data(1,:)));
                     start = [row(1), col(1)];
                     stop = [row(length(row(:))), col(length(col(:)))];
                     %insert data into mat
                     Mat(2:(stop(1)-start(1)+2), 1) = el(start(1):stop(1));
                     Mat(1, 2:(stop(2)-start(2)+2)) = az(start(2):stop(2));
                     Mat(1,1) = NaN;
                     Mat(2:(stop(1)-start(1)+2), 2:(stop(2)-start(2)+2)) = data((start(1)+1):(stop(1)+1), (start(2)+1):(stop(2)+1));
                end
%                 %find all the non-NaN data and it's indices
%                  [row,col] = find( ~isnan(data((2:length(data(:,1))),(2:length(data(1,:))))));
%                  %marking the 2 indices that limit our non-NaN data
%                  el = data(2:length(data(:,1)), 1);
%                  az = data(1, 2:length(data(1,:)));
%                  start = [row(1), col(1)];
%                  stop = [row(length(row(:))), col(length(col(:)))];
%                  %insert data into mat
%                  Mat(2:(stop(1)-start(1)+2), 1) = el(start(1):stop(1));
%                  Mat(1, 2:(stop(2)-start(2)+2)) = az(start(2):stop(2));
%                  Mat(1,1) = NaN;
%                  Mat(2:(stop(1)-start(1)+2), 2:(stop(2)-start(2)+2)) = data((start(1)+1):(stop(1)+1), (start(2)+1):(stop(2)+1));
            end
        end
    end
    
end%skip == 0
end