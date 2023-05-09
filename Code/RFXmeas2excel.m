function [] = RFXmeas2excel(Freq, Mode, SwitchAzEl, InputFileAdd, OutputFileAdd)
%Freq - frequency in this simulation
%Mode - anables choice between: Absolute gain, Absolute theta, Absolute phi,
%       Phase theta, Phase phi, Right hand circular polarization, Left hand 
%       circular polarization, Axial ratio  
%SwitchAzEl = switches between the azimuth and elevation
%InputFileAdd - input file address (path + name)
%OutputFilePath - output file path
%OutputFileName - output file name

%% Ensuring all inputs are valid
skip = 0; %instead of "break" function
if isempty(Freq)
    fprintf('Error: Please enter simulated frequency\n')
    skip = 1; %if skip changes to one the whole function will break
end
if isempty(Mode)
    fprintf('Error: Please enter requested FarField mode\n')
    skip = 1; %if skip changes to one the whole function will break
end
if isempty(SwitchAzEl)
    fprintf('Would you like to switch between azimuth and elevation?\nYes - Enter 1\nNo - enter 0')
    skip = 1; %if skip changes to one the whole function will break
end
if isempty(InputFileAdd)
    fprintf('Error: Please enter input file address (including path and name)\n')
    skip = 2; %special case! need it for the easter egg later on...
end
if isempty(OutputFileAdd)
    fprintf('Error: Please enter output file address\n')
    skip = 1; %if skip changes to one the whole function will break
end

if (skip == 0)
    %% importation
    
    %sheet 1 - Absolute gain
    if (Mode == 1)
        mode = 'Absolute Gain';
    end
    %sheet 2 - Abs(Theta)[dBi]
    if (Mode == 2)
        mode = 'Absolute Theta';
    end
    %sheet 3 - Abs(Phi)[dBi]
    if (Mode == 3)
        mode = 'Absolute Phi';
    end
    %sheet 4 - Phase(Theta)[deg.]
    if (Mode == 4)
        mode = 'Phase Theta';
    end
    %sheet 5 - Phase(Phi)[deg.]
    if (Mode == 5)
        mode = 'Phase Phi';
    end
    %sheet 6 - RHCP[dBi]
    if (Mode == 6)
        mode = 'RHCP';
    end
    %sheet 7 - LHCP[dBi]
    if (Mode == 7)
        mode = 'LHCP';
    end
    %sheet 8 - Ax.Ratio[dB]
    if (Mode == 8)
        mode = 'Axial Ratio';
    end
    
    %importation
    raw_data = xlsread(InputFileAdd, Mode);
    if(SwitchAzEl == 1)
        temp = raw_data(:,:)';
        clear raw_data;
        raw_data(:,:) = temp(:,:); 
    end
    
    %% extraction of elevation and azimuth vectors
    az = raw_data(1,2:length(raw_data(1,:)));
    el = raw_data(2:length(raw_data(:,1)),1);
    
    %calculating resolutions of az & el
    az_res = (round(10*abs(az(1)-az(2))))/10;
    el_res = (round(10*abs(el(1)-el(2))))/10;

    %creating NaN matrix
    data = NaN(round(360/el_res)+1, round((180/az_res)+2));
    
    %injecting az & el to data matrix
    data(1,2:length(raw_data(1,:))) = az(1):az_res:(180+az(1)+1);
    data(2:length(data(:,1)),1) = el(1):el_res:(360+el(1)-1);
    %injecting data
    data(2:length(raw_data(:,1)), 2:length(raw_data(1,:))) = raw_data(2:length(raw_data(:,1)), 2:length(raw_data(1,:)));
    %% printing matrices to excel files

        sheet_name = sprintf('%d GHz', Freq);
        if exist(OutputFileAdd,'file')
            xlswrite(OutputFileAdd,data(:,:),sheet_name);
        else
            %fill second sheet with azimuth and elevation data
            xlsrange_el = strcat('A2:A',num2str(length(data(2:length(data(:,1)), 1))+1));
            xlsrange_az = strcat('B2:B',num2str(length(data(1, 2:length(data(1,:))))+1));
            xlswrite(OutputFileAdd,cellstr('el'),'azimuth & elevation','A1:A1');
            xlswrite(OutputFileAdd,cellstr('az'),'azimuth & elevation','B1:B1');
            xlswrite(OutputFileAdd,cellstr('Mode'),'azimuth & elevation','C1:C1');
            xlswrite(OutputFileAdd,cellstr(mode),'azimuth & elevation','C2:C2');
            xlswrite(OutputFileAdd, data(2:length(data(:,1)), 1),'azimuth & elevation',xlsrange_el);
            xlswrite(OutputFileAdd, (data(1,2:length(data(1,:))))','azimuth & elevation',xlsrange_az);
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
    end
end