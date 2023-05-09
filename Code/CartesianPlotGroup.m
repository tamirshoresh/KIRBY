function [] = CartesiaPlotGroup(Freq, el, az, data, ModeCartes, ModeContour, ModeBW, BW, Mode)
%this function plots the spherical and polar plots using the
%'patternCustom' function available in matlab 2019.
%in order to use patternCustom we must first convert our input matrix into
%a table

%Freq - current frequency
% el, az - elevation, azimuth
%data - matrix containing all the Far Field data
%ModeCartes - indicates activation of cartesian plot mode
%ModeContour - indicates activation of contour plot mode
%ModeBW - indicates activation of threshold (BW) plot mode
%BW - input threshold
%Mode - mode specified in excel file (absolute gain, theta phase, so on...)

%% Ensuring all inputs are valid

skip = 0; %instead of "break" function
if isempty(Freq)
    fprintf('Error: Please enter specified frequency\n')
    skip = 1; %if skip changes to one the whole function will break
end
if isempty(BW)
    fprintf('Error: Please enter threshold level [dB]\n')
    skip = 1; %if skip changes to one the whole function will break
end

if (skip == 0)
%% creating the plots
    
    %deciding the plot units:
    if(strcmp(Mode,'Absolute Gain') || strcmp(Mode,'Absolute Theta') || strcmp(Mode,'Absolute Phi') || strcmp(Mode,'RHCP') || strcmp(Mode,'LHCP'))
        units = 'dBi';
    elseif(strcmp(Mode,'Phase Theta') || strcmp(Mode,'Phase Phi'))
        units = 'deg.';
    else
        units = 'dB';
    end
    
    [Az, El] = meshgrid(az, el);

    %cartesian plot
    if(ModeCartes == 1)
        FigName = sprintf('3D Cartezian Plot - %d GHz', Freq);
        figure('Name', FigName);
        s = surf(Az, El, data);
        %now we make everything pretty...
        s.EdgeColor = 'none';
        colorbar;
        colormap(jet);
        CB = colorbar;
        title(CB, units);
        title(FigName);
        xlabel('Azimuth / Phi (deg.)');
        ylabel('Elevation / Theta (deg.)');
    end

    %contour plot
    if(ModeContour == 1)
        FigName = sprintf('Contour Map - %d GHz', Freq);
        figure('Name', FigName);
        contourf(Az, El, data);
        %now we make everything pretty...
        colorbar;
        colormap(jet);
        CB = colorbar;
        title(CB, units);
        title(FigName);
        xlabel('Azimuth / Phi (deg.)');
        ylabel('Elevation / Theta (deg.)');
    end

    %treshold plot
    if(ModeBW == 1)
        thMat(:,:) = zeros(size(data(:,:))); %threshold Matrix
        for i = 1:length(el)
            for j = 1:length(az)
                if (data(i,j) > (max(max(data(:,:))) - str2double(BW)))
                    thMat(i,j) = 1;
                end
            end
        end
        FigName = sprintf('threshold at %d dB BandWidth - %d GHz', str2double(BW), Freq);
        figure('Name', FigName);
        pcolor(az, el, thMat);
        colormap(jet);
        title(FigName);
        xlabel('Azimuth / Phi (deg.)');
        ylabel('Elevation / Theta (deg.)');

    end
end
end

