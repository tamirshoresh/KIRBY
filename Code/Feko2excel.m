function [] = Feko2excel(Freq, InputFileAdd, OutputFileAdd)
%Freq - frequency in this simulation
%Mode - anables choice between: Absolute gain, Absolute theta, Phase theta,
%       Absolute phi, Phase phi, Axial ratio 
%InputFileAdd - input file address (path + name)
%OutputFilePath - output file path
%OutputFileName - output file name
%% Ensuring all inputs are valid
skip = 0; %instead of "break" function
if isempty(Freq)
    fprintf('Error: Please enter simulated frequency\n')
    skip = 1; %if skip changes to one the whole function will break
end
if isempty(InputFileAdd)
    fprintf('Error: Please enter input file address (including path and name)\n')
    skip = 1; %if skip changes to one the whole function will break
end
if isempty(OutputFileAdd)
    fprintf('Error: Please enter output file address\n')
    skip = 1; %if skip changes to one the whole function will break
end


%%
if (skip == 0)
    %% importation
    
    FF_raw_data = importdata(InputFileAdd);
    %row 1 - Theta [deg.] (elevation)
    %row 2 - Phi   [deg.] (azimuth)
    %row 3 - data
    FF_data = FF_raw_data.data(:,1:3);

    %% seperating elevation and azimuth from file

    az(1) = FF_data(1,2); %initializing the first azimuth measurement from file
    flag = 1; %marks when we have finished one iteration of azimuth
    counter = 1;
    %the following loop counts one iteration of az making it into a non-repeating vector as it goes
    while (flag == 1)
        if(FF_data(counter+1,2) == FF_data(1,2))
            flag = 2;
        else
            counter = counter+1;
            az(counter) = FF_data(counter,2);
        end
    end

    %extracting elevation vector
    for i = 1:(length(FF_data(:,2))/(length(az)))
        el(i) = FF_data(i*length(az),1);
    end

    %calculating resolutions of az & el
    el_res = abs(el(1)-el(2));
    az_res = abs(az(1)-az(2));

    %creating NaN matrix
    data = NaN(round(180/el_res)+2, round((360/az_res)+2));
    %injecting az & el to data matrix
    data(1,2:length(data(1,:))) = az(1):az_res:(360+az(1));
    data(2:length(data(:,1)),1) = el(1):el_res:(180+el(1));
    %injecting reshapes FF_data to data matrix
    data(2:length(data(:,1)), 2:length(data(1,:))) = (reshape(FF_data(:,3), length(az), length(el)))';
    
    %% printing matrices to excel files
    
    mode = 'Absolute Gain';
    sheet_name = sprintf('%d GHz', Freq);
    if exist(OutputFileAdd,'file')
        xlswrite(OutputFileAdd,data(:,:),sheet_name);
    else
        %fill second sheet with azimuth and elevation data
        xlsrange_el = strcat('A2:A',num2str(length(el)+1));
        xlsrange_az = strcat('B2:B',num2str(length(az)+1));
        xlswrite(OutputFileAdd,cellstr('el'),'azimuth & elevation','A1:A1');
        xlswrite(OutputFileAdd,cellstr('az'),'azimuth & elevation','B1:B1');
        xlswrite(OutputFileAdd,cellstr('Mode'),'azimuth & elevation','C1:C1');
        xlswrite(OutputFileAdd, cellstr(mode),'azimuth & elevation','C2:C2');
        xlswrite(OutputFileAdd, el','azimuth & elevation',xlsrange_el);
        xlswrite(OutputFileAdd, az','azimuth & elevation',xlsrange_az);
        xlswrite(OutputFileAdd,data(:,:),sheet_name);
        
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
end %if (skip == 0)

end