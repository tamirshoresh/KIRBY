function [] = SphericalPlotGroup(Freq, el, az, data, ModeSphere, ModeConstEl, ModeConstAz, el_const, az_const)
%this function plots the spherical and polar plots using the
%'patternCustom' function available in matlab 2019.
%in order to use patternCustom we must first convert our input matrix into
%a table

%Freq - current frequency
% el, az - elevation, azimuth
%data - matrix containing all the Far Field data
%ModeSphere - indicates activation of spherical plot mode
%ModeConstEl - indicates activation of constant elevation plot mode
%el_const - input of constant elevation
%ModeConstAz - indicates activation of constant azimuth plot mode
%az_const - input of constant azimuth


%% Ensuring all inputs are valid

skip = 0; %instead of "break" function
if isempty(Freq)
    fprintf('Error: Please enter specified frequency\n')
    skip = 1; %if skip changes to one the whole function will break
end
if isempty(el_const)
    fprintf('Error: Please enter constant elevation cut\n')
    skip = 1; %if skip changes to one the whole function will break
end
if isempty(az_const)
    fprintf('Error: Please enter constant azimuth cut\n')
    skip = 1; %if skip changes to one the whole function will break
end


if(skip == 0)
%% reshaping the matrix to be compatible to patternCustom

        %azimuth:
        dataTable(:,1) = repmat(el', 1, length(az)); 
        %elevation:
        dataTable(:,2) = reshape(repmat(az, length(el), 1),[],1);
        %data:
        dataTable(:,3) = reshape(data, [], 1);

%% creating the plots

        %spherical plot
        if (ModeSphere ==1)
            FigName = sprintf('3D Spherical Plot - %d GHz', Freq);
            figure('Name', FigName);
            patternCustom(dataTable(:,3), dataTable(:,1), dataTable(:,2))
            title(FigName);
            axis vis3d;
        end

        %constant elevation plot
        if (ModeConstEl ==1)
            FigName = sprintf('Constant Elevation Plot - %d GHz', Freq);
            figure('Name', FigName);
            patternCustom(dataTable(:,3), dataTable(:,1), dataTable(:,2), 'CoordinateSystem','polar','Slice','theta','SliceValue',str2num(el_const));
            title(FigName);
            axis vis3d;
        end

        %constant azimuth plot
        if (ModeConstAz ==1)
            FigName = sprintf('Constant Azimuth Plot - %d GHz', Freq);
            figure('Name', FigName);
            patternCustom(dataTable(:,3), dataTable(:,1), dataTable(:,2), 'CoordinateSystem','polar','Slice','phi','SliceValue',str2num(az_const));
            title(FigName);
            axis vis3d;
        end
    end
end

