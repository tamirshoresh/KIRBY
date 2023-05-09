function [] = ECmeas2excel_temp(FreqStr, Mode, az_res, el_res, CurrEl, InputFileAdd, OutputFileAdd)
%FreqStr - string of wanted frequencies separated by semicolons
%Mode - anables choice between: Absolte gain, Phase
%az_res - azimuth resolution
%el_res - elevation resolution
%InputFileAdd - input file address (path + name)
%CurrEl - current elevation input
%OutputFilePath - output file path
%OutputFileName - output file name
%% Ensuring all inputs are valid
skip = 0; %instead of "break" function
if isempty(FreqStr)
    fprintf('Error: Please enter frequency list\n')
    skip = 1; %if skip changes to one the whole function will break
end
if isempty(Mode)
    fprintf('Error: Please enter mode\n')
    skip = 1; %if skip changes to one the whole function will break
end
if isempty(az_res)
    fprintf('Error: Please enter azimuth resolution of input file\n')
    skip = 1; %if skip changes to one the whole function will break
end
if isempty(el_res)
    fprintf('Error: Please enter elevation resolution of input file\n')
    skip = 1; %if skip changes to one the whole function will break
end
if isempty(InputFileAdd)
    fprintf('Error: Please enter input file address (including path and name)\n')
    skip = 1; %if skip changes to one the whole function will break
end
if isempty(CurrEl)
    fprintf('Error: Please enter elevation of input file\n')
    skip = 1; %if skip changes to one the whole function will break
end
if isempty(OutputFileAdd)
    fprintf('Error: Please enter output file address\n')
    skip = 1; %if skip changes to one the whole function will break
end

%%
if (skip == 0)
    %% seperating wanted freq. & creation of NaN matrices
   
    %find semicolon indices
    SC = strfind(FreqStr,';');
    f_list = (str2num(FreqStr))';
    %creating an empty matrix for every frequency
    data = NaN([length(f_list), round((180/el_res)+2),round((360/az_res))]);
    
    %% importation and finding frequencies in file

    %retriving the data from the input file address into "file"
    file = str2mat(extractFileText(InputFileAdd));
    
    f_ind = strfind(file, 'Frequency'); %marks the indices of the word 'Frequency'
    %now add the indices of the end of the file to make sure our index won't
    %reach further than that
    f_ind(length(f_ind)+1) = length(file(1,:));
    
    for i=1:(length(f_ind)-1)
        %creation of wanted frequency vector
        f(i) = str2double(file(1,(f_ind(i)+10):(f_ind(i)+15)));
        for j=1:length(f_list)
            %comparison between the requested frequecy list (f_list) and the 
            %frequencies now retrived from the file (f)
            if (f(i) == f_list(j)) 
                %copy indices of requested frequencies and indices of
                %following frequency. these will mark the top and bottom
                %limits you must copy
                indices(j) = f_ind(i);
                indices_next(j) = f_ind(i+1);
            end
        end
    end
    
    %% checking if output file exists and creating data matrix
    if exist(OutputFileAdd,'file')
        %get the names of all the sheets
        [ExistingName, ExistingSheet] = xlsfinfo(OutputFileAdd);
        %get the frequencies and data from sheet 2 and on (sheet 1 is azimuth & elevation)
        for i=2:length(ExistingSheet)
            findGHz(i-1) = cell2mat(strfind(ExistingSheet(i),'GHz'));
            %change format of sheet name
            ExistingSheetName(i-1) = cellstr(cell2mat(ExistingSheet(i)));
            %build a matrix of all frequencies
            f_exist(i-1) = str2num(ExistingSheetName{i-1}(1:(findGHz(i-1)-2)));
        end
        if isempty(findGHz)
            fprintf('Congratulations! You fucked shit up again!\nWhy dont you try entering an appropriate Excel file\n');
        else
            %if file is correct, look for over-lapping frequencies
            for i = 1:length(f_exist)
                %the indices of the over_lapping (OL) frequencies in f_list
                indOL(i) = find(f_list == f_exist(i));
            end
            %copying into data matrix
            for i = 1:length(f_list)
                %check if the f_list indices equals to the indOL (meaning that frequency exists)
                if isempty((find(indOL == i))) %this frequency does not exists
                    %withdrawing the data for one specific frequency f_list(i)
                    DataMat = str2num(file(1, (indices(i)+104):(indices_next(i)-3)));
                    %withdrawind only the elevation data from previous data and
                    %writing it into the full data matrix
                    if (Mode == 1) %get absolute gain
                        mode = 'Absolute Gain';
                        data(i, (length(data(1,:,1))/2)+(CurrEl/el_res)+1, 2:length(data(1,1,:))) = (DataMat(:,2))';
                    else %get phase
                        mode = 'Phase';
                        data(i, (length(data(1,:,1))/2)+(CurrEl/el_res)+1, 2:length(data(1,1,:))) = (DataMat(:,3))';
                    end
                else %frequency exists in output file!
                    %withdrawing the data for one specific frequency f_list(i)
                    DataMat = str2num(file(1, (indices(i)+104):(indices_next(i)-3)));
                    %get existing data
                    existingData = xlsread(OutputFileAdd, i+1);
                    data(i,2:length(data(1,:,1)),2:length(data(1,1,:))) = existingData(2:length(existingData(:,1)),2:length(existingData(1,:)));
                    %withdrawind only the elevation data from previous data and
                    %writing it into the full data matrix
                    if (Mode == 1) %get absolute gain
                        mode = 'Absolute Gain';
                        data(i, (length(data(1,:,1))/2)+(CurrEl/el_res)+1, 2:length(data(1,1,:))) = (DataMat(:,2))';
                    else %get phase
                        mode = 'Phase';
                        data(i, (length(data(1,:,1))/2)+(CurrEl/el_res)+1, 2:length(data(1,1,:))) = (DataMat(:,3))';
                    end
                end
            end
        end
    else
        %OutputFile does not exist. Enter data into empty matrix:
        for i=1:length(f_list)
            %withdrawing the data for one specific frequency f_list(i)
            DataMat = str2num(file(1, (indices(i)+104):(indices_next(i)-3)));
            %withdrawind only the elevation data from previous data and
            %writing it into the full data matrix
            if (Mode == 1) %get absolute gain
                mode = 'Absolute Gain';
                data(i, (length(data(1,:,1))/2)+(CurrEl/el_res)+1, 2:length(data(1,1,:))) = (DataMat(:,2))';
            else %get phase
                mode = 'Phase';
                data(i, (length(data(1,:,1))/2)+(CurrEl/el_res)+1, 2:length(data(1,1,:))) = (DataMat(:,3))';
            end
        end
    end
    
    %% printing matrices to excel files
    
    for i = 1:length(f_list)
        data(i,2:length(data(:,:,1)),1) = -90:el_res:90;
        data(i,1,2:length(data(:,1,:))) = -179:az_res:179;
        sheet_name = sprintf('%d GHz', f_list(i));
        
        if exist(OutputFileAdd,'file')
            xlswrite(OutputFileAdd,squeeze(data(i,:,:)),sheet_name);
        else
            %fill second sheet with azimuth and elevation data
            el = squeeze(data(1,2:length(data(1,:,1)),1));
            az = squeeze(data(1,1,2:length(data(1,1,:))));
            xlsrange_el = strcat('A2:A',num2str(length(el)+1));
            xlsrange_az = strcat('B2:B',num2str(length(az)+1));
            xlswrite(OutputFileAdd,cellstr('el'),'azimuth & elevation','A1:A1');
            xlswrite(OutputFileAdd,cellstr('az'),'azimuth & elevation','B1:B1');
            xlswrite(OutputFileAdd,cellstr('Mode'),'azimuth & elevation','C1:C1');
            xlswrite(OutputFileAdd,cellstr(mode),'azimuth & elevation','C2:C2');
            xlswrite(OutputFileAdd, el','azimuth & elevation',xlsrange_el);
            xlswrite(OutputFileAdd, az,'azimuth & elevation',xlsrange_az);
            xlswrite(OutputFileAdd,squeeze(data(i,:,:)),sheet_name);

            %delete fist sheet
            %open excel file:
            objExcel = actxserver('Excel.Application');
            objExcel.Workbooks.Open(OutputFileAdd);
            %delete 'Sheet1'
            objExcel.ActiveWorkbook.Worksheets.Item('Sheet1').Delete;
            %save and close
            objExcel.ActiveWorkbook.Save;
            objExcel.ActiveWorkbook.Close;
            objExcel.Quit;
            objExcel.delete;
        end
    end
    
    
end %skip == 0
end