clear
close all
clc

%Inputs to calculate capacitance per mass
%==========================================================================
% ScanRate = [0.2, 0.1, 0.05, 0.02, 0.01, 0.005, 0.002];
ScanRate = 0.2;
Voltage = 1;
mass = 0.00326;
cycle = 3;
%==========================================================================

CapacitancePerMass = zeros(1, length(ScanRate));
legendTitles = zeros(length(ScanRate),1);
j = 1;
for iterator = ScanRate

    filenamea = [num2str(iterator), '.xlsx'];
    file = importdata(filenamea);
    data = getfield(file,'data');
    
    found = false;
    flag = 0;
    k1 = 1;
    for k2 = 1:length(data(:,1))
        
        while data(k2,5) == cycle
    
            found = true;
            
            xa(k1,1) = data(k2,1);
            ya(k1,1) = data(k2,3);

            k1 = k1+1;
            k2 = k2+1;
            
            
            if (k2 == length(data(:,1))+1)
               
                flag = 1;
                break;
            
            elseif (data(k2,5) ~= cycle)
               
                flag = 1;
                break;
                
            end
            
            

        end
        
        if (k2 == length(data(:,1))+1 && flag == 1)
               
                flag = 1;
                break;
        
        elseif ( data(k2,5) ~= cycle && flag == 1)
               
            break;
                
        end
        
    end

    
    if (found == false)
       
        disp("Such cycle does not exist !");
        
    end
    

    %%
    %%CV-20

    na = length(ya);
    % plot(abs(fft(y)))
    n_windowa = 3; % must be odd
    y_avga = zeros(na,1);
    %average 
    n_mida = round(na / 2) ;
    n_wa = (n_windowa-1)/2;
    for i = 1:n_mida
        if i == 1
            y_avga(i) = ya(i);
        elseif i <= n_wa 
           y_avga(i) = mean( ya(1:i+n_wa)); 
       elseif i >= (n_mida - n_wa)
          y_avga(i) = mean( ya(i-n_wa:n_mida));  
       else
          y_avga(i) = mean( ya(i-n_wa:i+n_wa)); 
       end
    end

    for i = n_mida+1:na
        if i == na
            y_avga(i) = ya(i);
        elseif i <= n_mida + n_wa 
           y_avga(i) = mean( ya(n_mida+1:i+n_wa)); 
       elseif i >= (na - (n_windowa-1)/2)
          y_avga(i) = mean( ya(i-n_wa:na));  
       else
          y_avga(i) = mean( ya(i-n_wa:i+n_wa)); 
       end
    end
    %% integral
    ds1a = 0;
    ds2a = 0;
    for i= 1:n_mida-1
        ds1a= ds1a + ( ( y_avga(i+1) + y_avga(i) )/2 * ( xa(i+1) - xa(i)) );
    end

    for i= n_mida+1:na-1
        ds2a= ds2a + ( ( y_avga(i+1) + y_avga(i) )/2 * ( xa(i) - xa(i+1)) );
    end
    Sa = ( ds1a - ds2a );

    
    C = Sa / (iterator*Voltage);
    CapacitancePerMass(j) = C/mass;

    % figure('Name','CV')
    figure(j);
    fig = plot(xa,y_avga,'r');
    hold on
    title('CV-Curve');
    xlabel('Voltage(V)')
    ylabel('Current(A)')
    grid on
    legend(['Scan Rate = ', num2str(iterator), ' (V/s))  &  Capacitance ', num2str(CapacitancePerMass(j)),' (F/g)'], 'location', 'northwest');

    
    savefig([num2str(iterator), '.fig']);
    saveas(fig, [num2str(iterator), '.bmp']);
    
    
    
    figure(length(ScanRate)+1);
    fig2 = plot(xa,y_avga);
    hold on
    title('CV-Curve');
    xlabel('Voltage(V)')
    ylabel('Current(A)')
    grid on
    
    legendTitles(j,1) = iterator;

    
    j = j+1;
    
end

legend(num2str(legendTitles), 'location', 'northwest');
savefig('AllTogether.fig');
saveas(fig2,'AllTogether.bmp');



rngeCapacitance = num2str(length(CapacitancePerMass)+1,'B1:B%d');
rngeScanRate = num2str(length(ScanRate)+1,'A1:A%d');

xlswrite ('Capacitances.xlsx',CapacitancePerMass(:),1,rngeCapacitance);
xlswrite ('Capacitances.xlsx',ScanRate(:),1,rngeScanRate);