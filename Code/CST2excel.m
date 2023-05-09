function [] = CST2excel(Freq, Mode, InputFileAdd, OutputFileAdd)
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
if isempty(Mode)
    fprintf('Error: Please enter requested FarField mode\n')
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

%%
if (skip ~= 2)
    %% importation
    
    FF_raw_data = importdata(InputFileAdd);
    %row 1 - Theta [deg.] (elevation)
    %row 2 - Phi   [deg.] (azimuth)
    %row 3 - Abs(Gain)[dBi   ]
    %row 4 - Abs(Theta)[dBi   ]
    %row 5 - Phase(Theta)[deg.]
    %row 6 - Abs(Phi  )[dBi   ]
    %row 7 - Phase(Phi  )[deg.]
    %row 8 - Ax.Ratio[dB    ]
    
    switch Mode
        case 1 %Absolute gain
            FF_data = FF_raw_data.data(:,1:2);
            FF_data(:,3) = FF_raw_data.data(:,3);
            mode = 'Absolute Gain';
        case 2 %Absolute theta
            FF_data = FF_raw_data.data(:,1:2);
            FF_data(:,3) = FF_raw_data.data(:,4);
            mode = 'Absolute Theta';
        case 3 %Phase theta
            FF_data = FF_raw_data.data(:,1:2);
            FF_data(:,3) = FF_raw_data.data(:,5);
            mode = 'Phase Theta';
        case 4 %Absolute phi
            FF_data = FF_raw_data.data(:,1:2);
            FF_data(:,3) = FF_raw_data.data(:,6);
            mode = 'Absolute Phi';
        case 5 %Phase phi
            FF_data = FF_raw_data.data(:,1:2);
            FF_data(:,3) = FF_raw_data.data(:,7);
            mode = 'Phase Phi';
        case 6 %Axial ratio
            FF_data = FF_raw_data.data(:,1:2);
            FF_data(:,3) = FF_raw_data.data(:,8);
            mode = 'Axial Ratio';
        otherwise
            fprintf('you know, youre not supposed to do that...\n')    
            skip = 1; %if skip changes to one the whole function will break
    end

    if(skip == 0)
        %% seperating elevation and azimuth from file

        el(1) = FF_data(1,1); %initializing the first azimuth measurement from file
        flag = 1; %marks when we have finished one iteration of azimuth
        counter = 1;
        %the following loop counts one iteration of az making it into a non-repeating vector as it goes
        while (flag == 1)
            if(FF_data(counter+1,1) == FF_data(1,1))
                flag = 2;
            else
                counter = counter+1;
                el(counter) = FF_data(counter,1);
            end
        end

        %extracting elevation vector
        for i = 1:(length(FF_data(:,2))/(length(el)))
            az(i) = FF_data(i*length(el),2);
        end

        %calculating resolutions of az & el
        el_res = abs(el(1)-el(2));
        az_res = abs(az(1)-az(2));

        %creating NaN matrix
        data = NaN(round(360/el_res)+1, round((180/az_res)+2));
        %injecting az & el to data matrix
        data(1,2:length(data(1,:))) = az(1):az_res:(180+az(1));
        data(2:length(data(:,1)),1) = el(1):el_res:(360+el(1)-1);
        %injecting reshapes FF_data to data matrix
        data(2:length(data(:,1)), 2:length(data(1,:))) = reshape(FF_data(:,3), length(el), length(az));

        %% printing matrices to excel files

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
            xlswrite(OutputFileAdd,cellstr(mode),'azimuth & elevation','C2:C2');
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
end %if (skip ~= 2)

end