% Authors: Alissa P., Henry T. 
% this script reads microglia parameters and does statistical testing
% between PEN and SHAM ROIs 
clc 
close all 
clear all

%% importing xcel data 
[data,sheet_names] = xcel_file_read('C:\Users\Henry\MATLAB\Mourad Lab\AD_particle_analysis\Data\Microglia\Chi and sham microglia parameters.xlsx') ;

% T = readtable('AB Plaque Quantification (particle analysis) CENTER ROI.xlsx');

%% assigning xcel sheet microglia parameters to variables 
% data{k}(row_index, col_index) key for calling items in struct 

% Column data in each sheet: 
% CM1A CM1B	CM2A CM2B CM3A CM3B	NaN	SM1A SM1B SM2A SM2B	SM3A SM3B

% ROI A 
for i = 1:6 
    penROIA.(sheet_names{1,i}) = [data{i}(:, 1) data{i}(:, 3) data{i}(:, 5)]; % data is in 3 columns 
    shamROIA.(sheet_names{1,i}) = [data{i}(:, 8) data{i}(:, 10) data{i}(:, 12)];
    % converting matrixes to vectors 
    penROIA.(sheet_names{1,i}) = penROIA.(sheet_names{1,i})(:);
    shamROIA.(sheet_names{1,i}) = shamROIA.(sheet_names{1,i})(:);
end 

% ROI B 
for i = 1:6 
    penROIB.(sheet_names{1,i}) = [data{2}(:, 1) data{i}(:, 4) data{i}(:, 6)];
    shamROIB.(sheet_names{1,i}) = [data{i}(:, 9) data{i}(:, 11) data{i}(:, 13)];
    % converting matrixes to vectors 
    penROIB.(sheet_names{1,i}) = penROIB.(sheet_names{1,i})(:);
    shamROIB.(sheet_names{1,i}) = shamROIB.(sheet_names{1,i})(:);
end 


%% statisical analysis 

% QQ ROIA 
figure(1)
ROIA_AR = qqplot(penROIA.AR,shamROIA.AR) ;
        title('ROIA QQ plot of chikodi vs sham AR')
        xlabel('Quantiles of chikodi data')
        ylabel('Quantiles of sham data')
%         hold on 
%         plot(0:100:16000, 0:100:16000) 
    MW_ROIA_AR = ranksum(penROIA.AR,shamROIA.AR);
    	% boxplot 
        figure(2)
        boxplot([penROIA.AR,shamROIA.AR],'Notch','on','Labels', {'ROIA Chi AR', 'ROIA Sham AR'}) 
        hold on 
        %scatterplot % this takes sooooo long
%         scatter(ones(length(penROIA.AR)),penROIA.AR,3,'r','filled')
%         hold on 
%         scatter(2.*ones(length(shamROIA.AR)),shamROIA.AR,3,'r','filled')
        title('ROIA Chi vs SHAM AR') 
        ylabel('idk what the AR values mean') 
        
figure(3)
ROIA_Solidity = qqplot(penROIA.Solidity,shamROIA.Solidity) ;
        title('ROIA QQ plot of chikodi vs sham Solidity')
        xlabel('Quantiles of chikodi data')
        ylabel('Quantiles of sham data')
        MW_ROIA_Solidity = ranksum(penROIA.Solidity,shamROIA.Solidity);
       % boxplot 
        figure(4)
        boxplot([penROIA.Solidity,shamROIA.Solidity],'Notch','on','Labels', {'ROIA Chi Solidity', 'ROIA Sham Solidity'}) 
        title('ROIA Chi vs SHAM Solidity') 
        ylabel('idk what the Solidity values mean') 
        
figure(5)
ROIA_Feret = qqplot(penROIA.Feret,shamROIA.Feret) ;
        title('ROIA QQ plot of chikodi vs sham Feret')
        xlabel('Quantiles of chikodi data')
        ylabel('Quantiles of sham data')
        MW_ROIA_Feret = ranksum(penROIA.Feret,shamROIA.Feret);
         % boxplot 
        figure(6)
        boxplot([penROIA.Feret,shamROIA.Feret],'Notch','on','Labels', {'ROIA Chi Feret', 'ROIA Sham Feret'}) 
        title('ROIA Chi vs SHAM Feret') 
        ylabel('idk what the feret y axis units are') 
        
figure(7)
ROIA_Perimeter = qqplot(penROIA.Perimeter,shamROIA.Perimeter) ;
        title('ROIA QQ plot of chikodi vs sham Perimeter')
        xlabel('Quantiles of chikodi data')
        ylabel('Quantiles of sham data')
        MW_ROIA_Perimeter = ranksum(penROIA.Perimeter,shamROIA.Perimeter);
       % boxplot 
        figure(8)
        boxplot([penROIA.Perimeter,shamROIA.Perimeter],'Notch','on','Labels', {'ROIA Chi Perimeter', 'ROIA Sham Perimeter'}) 
        title('ROIA Chi vs SHAM Perimeter') 
        ylabel('Perimeter units')
        
figure(9)
ROIA_Area = qqplot(penROIA.Area,shamROIA.Area) ;
        title('ROIA QQ plot of chikodi vs sham Area')
        xlabel('Quantiles of chikodi data')
        ylabel('Quantiles of sham data')
        MW_ROIA_Area = ranksum(penROIA.Area,shamROIA.Area);
        % boxplot 
        figure(10)
        boxplot([penROIA.Area,shamROIA.Area],'Notch','on','Labels', {'ROIA Chi Areas', 'ROIA Sham Areas'}) 
        title('ROIA Chi vs SHAM Areas') 
        ylabel('Area (um^2?)')
        
figure(11)
ROIA_Roundness = qqplot(penROIA.Roundness,shamROIA.Roundness) ;
        title('ROIA QQ plot of chikodi vs sham Roundness')
        xlabel('Quantiles of chikodi data')
        ylabel('Quantiles of sham data')
        MW_ROIA_Roundness = ranksum(penROIA.Roundness,shamROIA.Roundness);
       % boxplot 
        figure(10)
        boxplot([penROIA.Roundness,shamROIA.Roundness],'Notch','on','Labels', {'ROIA Chi Roundness', 'ROIA Sham Roundness'}) 
        title('ROIA Chi vs SHAM Roundness') 
        ylabel('Roundness units?')
        
%%  ROIB QQS 

figure(12)
ROIB_AR = qqplot(penROIB.AR,shamROIB.AR) ;
        title('ROIB QQ plot of chikodi vs sham AR')
        xlabel('Quantiles of chikodi data')
        ylabel('Quantiles of sham data')
%         hold on 
%         plot(0:100:16000, 0:100:16000) 
        MW_ROIB_AR = ranksum(penROIB.AR,shamROIB.AR);
        % boxplot 
        figure(13)
        boxplot([penROIB.AR,shamROIB.AR],'Notch','on','Labels', {'ROIB Chi AR', 'ROIB Sham AR'}) 
        title('ROIB Chi vs SHAM AR') 
        ylabel('AR units?')
        
figure(14)
ROIB_Solidity = qqplot(penROIB.Solidity,shamROIB.Solidity) ;
        title('ROIB QQ plot of chikodi vs sham Solidity')
        xlabel('Quantiles of chikodi data')
        ylabel('Quantiles of sham data')
        MW_ROIB_Solidity = ranksum(penROIB.Solidity,shamROIB.Solidity);
        % boxplot 
        figure(15)
        boxplot([penROIB.Solidity,shamROIB.Solidity],'Notch','on','Labels', {'ROIB Chi Solidity', 'ROIB Sham Solidity'}) 
        title('ROIB Chi vs SHAM Solidity') 
        ylabel('Solidity units?')
        
figure(16)
ROIB_Feret = qqplot(penROIB.Feret,shamROIB.Feret) ;
        title('ROIB QQ plot of chikodi vs sham Feret')
        xlabel('Quantiles of chikodi data')
        ylabel('Quantiles of sham data')
        MW_ROIB_Feret = ranksum(penROIB.Feret,shamROIA.Feret);
        % boxplot 
        figure(17)
        boxplot([penROIB.Feret,shamROIB.Feret],'Notch','on','Labels', {'ROIB Chi Feret', 'ROIB Sham Feret'}) 
        title('ROIB Chi vs SHAM Feret') 
        ylabel('Feret units?')
        
figure(18)
ROIB_Perimeter = qqplot(penROIB.Perimeter,shamROIB.Perimeter) ;
        title('ROIB QQ plot of chikodi vs sham Perimeter')
        xlabel('Quantiles of chikodi data')
        ylabel('Quantiles of sham data')
        MW_ROIB_Perimeter = ranksum(penROIB.Perimeter,shamROIB.Perimeter);
        % boxplot 
        figure(19)
        boxplot([penROIB.Perimeter,shamROIB.Perimeter],'Notch','on','Labels', {'ROIB Chi Perimeter', 'ROIB Sham Perimeter'}) 
        title('ROIB Chi vs SHAM Perimeter') 
        ylabel('Perimeter units?')
        
figure(20)
ROIB_Area = qqplot(penROIB.Area,shamROIB.Area) ;
        title('ROIB QQ plot of chikodi vs sham Area')
        xlabel('Quantiles of chikodi data')
        ylabel('Quantiles of sham data')
        MW_ROIB_Area = ranksum(penROIB.Area,shamROIB.Area);
        % boxplot 
        figure(21)
        boxplot([penROIB.Area,shamROIB.Area],'Notch','on','Labels', {'ROIB Chi Area', 'ROIB Sham Area'}) 
        title('ROIB Chi vs SHAM Area') 
        ylabel('Area (um^2?)')

figure(21)
ROIB_Roundness = qqplot(penROIB.Roundness,shamROIB.Roundness) ;
        title('ROIB QQ plot of chikodi vs sham Roundness')
        xlabel('Quantiles of chikodi data')
        ylabel('Quantiles of sham data')
        MW_ROIB_Roundness = ranksum(penROIB.Roundness,shamROIB.Roundness);
        % boxplot 
        figure(22)
        boxplot([penROIB.Roundness,shamROIB.Roundness],'Notch','on','Labels', {'ROIB Chi Roundness', 'ROIB Sham Roundness'}) 
        title('ROIB Chi vs SHAM Roundness') 
        ylabel('Roundness units?')
        
%% Creating new excel sheet with ROIA and ROIB Parameter Arrays 
% CM1A CM1B	CM2A CM2B CM3A CM3B	NaN	SM1A SM1B SM2A SM2B	SM3A SM3B
    filename = 'Microglia Parameter Particle Analysis.xlsx';
    
% pen ROIA 
    writematrix(penROIA.AR,filename,'Sheet','C_ROIA','Range','A2')
    writematrix(penROIA.Solidity,filename,'Sheet','C_ROIA','Range','B2')
    writematrix(penROIA.Feret,filename,'Sheet','C_ROIA','Range','C2')
    writematrix(penROIA.Perimeter,filename,'Sheet','C_ROIA','Range','D2')
    writematrix(penROIA.Area,filename,'Sheet','C_ROIA','Range','E2')
    writematrix(penROIA.Roundness,filename,'Sheet','C_ROIA','Range','F2')
    
    % titling 
    A1title = {'AR'};
    writecell(A1title,filename,'Sheet','C_ROIA','Range','A1');
    B1title = {'Solidity'};
    writecell(B1title,filename,'Sheet','C_ROIA','Range','B1');
    C1title = {'Feret'};
    writecell(C1title,filename,'Sheet','C_ROIA','Range','C1');
    D1title = {'Perimeter'};
    writecell(D1title,filename,'Sheet','C_ROIA','Range','D1');
    E1title = {'Area'};
    writecell(E1title,filename,'Sheet','C_ROIA','Range','E1');
    F1title = {'Roundness'};
    writecell(F1title,filename,'Sheet','C_ROIA','Range','F1');
    
% pen ROIB 
    
    writematrix(penROIB.AR,filename,'Sheet','C_ROIB','Range','A2')
    writematrix(penROIB.Solidity,filename,'Sheet','C_ROIB','Range','B2')
    writematrix(penROIB.Feret,filename,'Sheet','C_ROIB','Range','C2')
    writematrix(penROIB.Perimeter,filename,'Sheet','C_ROIB','Range','D2')
    writematrix(penROIB.Area,filename,'Sheet','C_ROIB','Range','E2')
    writematrix(penROIB.Roundness,filename,'Sheet','C_ROIB','Range','F2')
    
    % titling 
    A1title = {'AR'};
    writecell(A1title,filename,'Sheet','C_ROIB','Range','A1');
    B1title = {'Solidity'};
    writecell(B1title,filename,'Sheet','C_ROIB','Range','B1');
    C1title = {'Feret'};
    writecell(C1title,filename,'Sheet','C_ROIB','Range','C1');
    D1title = {'Perimeter'};
    writecell(D1title,filename,'Sheet','C_ROIB','Range','D1');
    E1title = {'Area'};
    writecell(E1title,filename,'Sheet','C_ROIB','Range','E1');
    F1title = {'Roundness'};
    writecell(F1title,filename,'Sheet','C_ROIB','Range','F1');
    
% SHAM ROIA 
    
    writematrix(shamROIA.AR,filename,'Sheet','S_ROIA','Range','A2')
    writematrix(shamROIA.Solidity,filename,'Sheet','S_ROIA','Range','B2')
    writematrix(shamROIA.Feret,filename,'Sheet','S_ROIA','Range','C2')
    writematrix(shamROIA.Perimeter,filename,'Sheet','S_ROIA','Range','D2')
    writematrix(shamROIA.Area,filename,'Sheet','S_ROIA','Range','E2')
    writematrix(shamROIA.Roundness,filename,'Sheet','S_ROIA','Range','F2')
    
    % titling 
    writecell(A1title,filename,'Sheet','S_ROIA','Range','A1');
    writecell(B1title,filename,'Sheet','S_ROIA','Range','B1');
    writecell(C1title,filename,'Sheet','S_ROIA','Range','C1');
    writecell(D1title,filename,'Sheet','S_ROIA','Range','D1');
    writecell(E1title,filename,'Sheet','S_ROIA','Range','E1');
    writecell(F1title,filename,'Sheet','S_ROIA','Range','F1');
    
% SHAM ROIB 

    writematrix(shamROIB.AR,filename,'Sheet','S_ROIB','Range','A2')
    writematrix(shamROIB.Solidity,filename,'Sheet','S_ROIB','Range','B2')
    writematrix(shamROIB.Feret,filename,'Sheet','S_ROIB','Range','C2')
    writematrix(shamROIB.Perimeter,filename,'Sheet','S_ROIB','Range','D2')
    writematrix(shamROIB.Area,filename,'Sheet','S_ROIB','Range','E2')
    writematrix(shamROIB.Roundness,filename,'Sheet','S_ROIB','Range','F2')
    
    % titling 
    writecell(A1title,filename,'Sheet','S_ROIB','Range','A1');
    writecell(B1title,filename,'Sheet','S_ROIB','Range','B1');
    writecell(C1title,filename,'Sheet','S_ROIB','Range','C1');
    writecell(D1title,filename,'Sheet','S_ROIB','Range','D1');
    writecell(E1title,filename,'Sheet','S_ROIB','Range','E1');
    writecell(F1title,filename,'Sheet','S_ROIB','Range','F1');
    
% MWs and summary sheet 
% ROIA
    writematrix(MW_ROIA_AR,filename,'Sheet','Summary','Range','B2')
    writematrix(MW_ROIA_Solidity,filename,'Sheet','Summary','Range','C2')
    writematrix(MW_ROIA_Feret,filename,'Sheet','Summary','Range','D2')
    writematrix(MW_ROIA_Perimeter,filename,'Sheet','Summary','Range','E2')
    writematrix(MW_ROIA_Area,filename,'Sheet','Summary','Range','F2')
    writematrix(MW_ROIA_Roundness,filename,'Sheet','Summary','Range','G2')
    writecell({'MW_ROIA_AR'},filename,'Sheet','Summary','Range','B1');
    writecell({'MW_ROIA_Solidity,'},filename,'Sheet','Summary','Range','C1');
    writecell({'MW_ROIA_Feret'},filename,'Sheet','Summary','Range','D1');
    writecell({'MW_ROIA_Perimeter'},filename,'Sheet','Summary','Range','E1');
    writecell({'MW_ROIA_Area'},filename,'Sheet','Summary','Range','F1');
    writecell({'MW_ROIA_Roundness'},filename,'Sheet','Summary','Range','G1');
   
% ROIB 
    writematrix(MW_ROIB_AR,filename,'Sheet','Summary','Range','G2')
    writematrix(MW_ROIB_Solidity,filename,'Sheet','Summary','Range','H2')
    writematrix(MW_ROIB_Feret,filename,'Sheet','Summary','Range','I2')
    writematrix(MW_ROIB_Perimeter,filename,'Sheet','Summary','Range','J2')
    writematrix(MW_ROIB_Area,filename,'Sheet','Summary','Range','K2')
    writematrix(MW_ROIB_Roundness,filename,'Sheet','Summary','Range','L2')
    writecell({'MW_ROIB_AR'},filename,'Sheet','Summary','Range','G1');
    writecell({'MW_ROIB_Solidity,'},filename,'Sheet','Summary','Range','H1');
    writecell({'MW_ROIB_Feret'},filename,'Sheet','Summary','Range','I1');
    writecell({'MW_ROIB_Perimeter'},filename,'Sheet','Summary','Range','J1');
    writecell({'MW_ROIB_Area'},filename,'Sheet','Summary','Range','K1');
    writecell({'MW_ROIB_Roundness'},filename,'Sheet','Summary','Range','L1');
    
    writecell({'p values'},filename,'Sheet','Summary','Range','A2');
    
%     % medians and means 
%     writematrix(median_bob,filename,'Sheet','Summary','Range','B5')
%     writematrix(median_chi,filename,'Sheet','Summary','Range','C5')
%     writematrix(median_sham,filename,'Sheet','Summary','Range','D5')
%     
%     writematrix(mean_bob,filename,'Sheet','Summary','Range','B6')
%     writematrix(mean_chi,filename,'Sheet','Summary','Range','C6')
%     writematrix(mean_sham,filename,'Sheet','Summary','Range','D6')
%     
%     writecell({'bob'},filename,'Sheet','Summary','Range','B4');
%     writecell({'chi'},filename,'Sheet','Summary','Range','C4');
%     writecell({'sham'},filename,'Sheet','Summary','Range','D4');
%     writecell({'medians'},filename,'Sheet','Summary','Range','A5');
%     writecell({'means'},filename,'Sheet','Summary','Range','A6');