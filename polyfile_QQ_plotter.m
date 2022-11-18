% Authors: Alissa P., Henry T. 
% this script reads histology data from AD excel spreadsheets and produces QQ plots of plaque sizes for different US protocol cohorts
% it can also create a new consolidated spreadsheet from multiple files 
clc 
close all 
clear all

%% creating new xcel sheet decision
decision = input("Would you like to create a consolidated excel file for the AB plaques by US protocol? ('1'=yes, '0'=no): ") ;
decision2 =  input("Would you like to create a second sheet of filtered values? ('1'=yes, '0'=no): ") ;
whatdata = input("Run series A ('1'), series B ('2'), or 'FULL_' data('3')?: ") ;
%% reading xcel data 

if whatdata == 3 
    % FULL series chikodi data 
    file1 = xlsread('Data\FULL\FULL_Chik_M1_AB_ROI_1_ch00.tiff AB results table.csv');
    file2 = xlsread('Data\FULL\FULL_Chik_M1_AB_ROI_2_ch00.tiff AB results table.csv');
    file3 = xlsread('Data\FULL\FULL_Chik_M2_AB_ROI_1_ch00.tiff AB results table.csv');
    file4 = xlsread('Data\FULL\FULL_Chik_M2_AB_ROI_2_ch00.tiff AB results table.csv');
    file5 = xlsread('Data\FULL\FULL_Chik_M3_AB_ROI_1_ch00.tiff AB results table.csv');
    file6 = xlsread('Data\FULL\FULL_Chik_M3_AB_ROI_2_ch00.tiff AB results table.csv');
    plaques_chi = [file1(:,3) ; file2(:,3) ; file3(:,3); file4(:,3) ; file5(:,3) ; file6(:,3)]';

    % FULL series SHAM data 
    file7 = xlsread('Data\FULL\FULL_Sham_M1_AB_ROI_1_ch00.tiff AB results table.csv');
    file8 = xlsread('Data\FULL\FULL_Sham_M1_AB_ROI_2_ch00.tiff AB results table.csv');
    file9 = xlsread('Data\FULL\FULL_Sham_M2_AB_ROI_1 _ch00.tiff AB results table.csv');
    file10 = xlsread('Data\FULL\FULL_Sham_M2_AB_ROI_2_ch00.tiff AB results table.csv');
    file11 = xlsread('Data\FULL\FULL_Sham_M3_AB_ROI_1_ch00.tiff AB results table.csv');
    file12 = xlsread('Data\FULL\FULL_Sham_M3_AB_ROI_2_ch00.tiff AB results table.csv');
    plaques_sham = [file7(:,3) ; file8(:,3) ; file9(:,3); file10(:,3) ; file11(:,3) ; file12(:,3)]';
elseif whatdata == 1
    % bobola 
    file1 = xlsread('Data\Series_A\Bob_11_12_21 ABM Slice 11 LH ROI.tif Plaque Analysis.csv');
    file2 = xlsread('Data\Series_A\Bob_12_10_21_ABM_S5 LH_ROI.tif Plaque Analysis.csv');
    file3 = xlsread('Data\Series_A\Bob_12_18_21_ABM_S5 probably LH ROI.tif Plaque Analysis.csv');
    plaques_bob = [file1(:,2) ; file2(:,2) ; file3(:,2)]';
    % plaques_bob = (plaques_bob.*1000000)';

    % chikodi 
    file4 = xlsread('Data\Series_A\Chi_11_11_21_ABM_S5 LH ROI.tif Plaque Analysis.csv');
    file5 = xlsread('Data\Series_A\Chi_12_03_21 ABM_S5 LH_ROI.tif Plaque Analysis.csv');
    file6 = xlsread('Data\Series_A\Chi_12_09_21 ABM_S5 LH ROI.tif Plaque Analysis.csv');
    plaques_chi = [file4(:,2) ; file5(:,2) ; file6(:,2)]';
    % plaques_chi = (plaques_chi.*1000000)';

    % sham
    file7 = xlsread('Data\Series_A\SHAM 12.18.21 ABM S11 full LHf.tif (RGB).tif Plaque Analysis.csv');
    file8 = xlsread('Data\Series_A\Sham_M2_S3A_LH_fulltif.tif Plaque Analysis.csv');
    file9 = xlsread('Data\Series_A\Sham_M3_S6A LH_ROItif.tif Plaque Analysis.csv');
    plaques_sham = [file7(:,2) ; file8(:,2) ; file9(:,2)]';
    % sham_plaques = (sham_plaques.*1000000)'; % already in um^2 in sheet
end 

%% QQ plots 

if whatdata ~= 3 
    % chikodi vs. sham 
    figure(1)
    chi_vs_sham = qqplot(plaques_chi,plaques_sham) ;
            title('QQ plot of chi vs sham')
            xlabel('Quantiles of Chikodi data')
            ylabel('Quantiles of sham data')
            hold on 
            plot(0:100:16000, 0:100:16000) 
    % 
    % Bobola vs. sham 
    figure(2)
    bob_vs_sham = qqplot(plaques_bob,plaques_sham) ;
            title('QQ plot of bob vs sham')
            xlabel('Quantiles of bobola data')
            ylabel('Quantiles of sham data')
            hold on 
            plot(0:100:16000, 0:100:16000) 
    %         
    % % Bobola vs. chikodi 
    figure(3)
    bob_vs_chi = qqplot(plaques_bob,plaques_chi) ;
            title('QQ plot of bob vs chi')
            xlabel('Quantiles of bobola data')
            ylabel('Quantiles of chikodi data')
            hold on 
            plot(0:100:16000, 0:100:16000) 
else 
    % chikodi vs. sham 
    figure(1)
    chi_vs_sham = qqplot(plaques_chi,plaques_sham) ;
            title('QQ plot of chi vs sham')
            xlabel('Quantiles of Chikodi data')
            ylabel('Quantiles of sham data')
            hold on 
            plot(0:100:16000, 0:100:16000) 
end 

%% Mann Whitneys 

if whatdata ~= 3 
    MW_chi_vs_sham = ranksum(plaques_chi,plaques_sham);
    MW_bob_vs_sham = ranksum(plaques_bob,plaques_sham);
    MW_bob_vs_chi = ranksum(plaques_bob,plaques_chi);

    median_bob = median(plaques_bob);
    median_chi = median(plaques_chi);
    median_sham = median(plaques_sham);

    mean_bob = mean(plaques_bob);
    mean_chi = mean(plaques_chi);
    mean_sham = mean(plaques_sham);
else 
    MW_chi_vs_sham = ranksum(plaques_chi,plaques_sham);
    
    median_chi = median(plaques_chi);
    median_sham = median(plaques_sham);
    
    mean_chi = mean(plaques_chi);
    mean_sham = mean(plaques_sham);
end 

%% exporting data to xcel sheet 
if decision == 1 
    filename = 'Left Hemisphere Particle Analysis.xlsx';
    writematrix(plaques_bob',filename,'Sheet','Bobola (m=3)','Range','A2')
    writematrix(plaques_chi',filename,'Sheet','Chikodi (m=3)','Range','A2')
    writematrix(plaques_sham',filename,'Sheet','Sham (m=3)','Range','A2')
    % titling A1 column for each sheet
    A1title = {'plaque area (um^2)'};
    writecell(A1title,filename,'Sheet','Bobola (m=3)','Range','A1');
    writecell(A1title,filename,'Sheet','Chikodi (m=3)','Range','A1');
    writecell(A1title,filename,'Sheet','Sham (m=3)','Range','A1');
    
    % MWs and summary sheet 
    writematrix(MW_bob_vs_chi,filename,'Sheet','Summary','Range','B2')
    writematrix(MW_bob_vs_sham,filename,'Sheet','Summary','Range','C2')
    writematrix(MW_chi_vs_sham,filename,'Sheet','Summary','Range','D2')
    writecell({'MW bob vs. chi'},filename,'Sheet','Summary','Range','B1');
    writecell({'MW bob vs. sham,'},filename,'Sheet','Summary','Range','C1');
    writecell({'MW chi vs. sham'},filename,'Sheet','Summary','Range','D1');
    writecell({'p values'},filename,'Sheet','Summary','Range','A2');
    
    % medians and means 
    writematrix(median_bob,filename,'Sheet','Summary','Range','B5')
    writematrix(median_chi,filename,'Sheet','Summary','Range','C5')
    writematrix(median_sham,filename,'Sheet','Summary','Range','D5')
    
    writematrix(mean_bob,filename,'Sheet','Summary','Range','B6')
    writematrix(mean_chi,filename,'Sheet','Summary','Range','C6')
    writematrix(mean_sham,filename,'Sheet','Summary','Range','D6')
    
    writecell({'bob'},filename,'Sheet','Summary','Range','B4');
    writecell({'chi'},filename,'Sheet','Summary','Range','C4');
    writecell({'sham'},filename,'Sheet','Summary','Range','D4');
    writecell({'medians'},filename,'Sheet','Summary','Range','A5');
    writecell({'means'},filename,'Sheet','Summary','Range','A6');
end 

%% filtering plaques to different sizes + new excel sheet 

if decision2 == 1 
    threshold = input("What is the threshold value?: ") ;
    operation = string(input("What is the operation? (ie. 'greater than', 'less than', 'greater than or equal to' etc.): ")) ;
    get_filtered_xcel(plaques_bob, plaques_chi, plaques_sham, threshold, operation);


% plaques_bob = plaques_bob(plaques_bob >= 1000) ;
% plaques_chi = plaques_chi(plaques_chi >= 1000) ;
% plaques_sham = plaques_sham(plaques_sham >= 1000) ;
% 
% median_bob = median(plaques_bob);
% median_chi = median(plaques_chi);
% median_sham = median(plaques_sham);
% 
% mean_bob = mean(plaques_bob);
% mean_chi = mean(plaques_chi);
% mean_sham = mean(plaques_sham);
% 
% MW_chi_vs_sham = ranksum(plaques_chi,plaques_sham);
% MW_bob_vs_sham = ranksum(plaques_bob,plaques_sham);
% MW_bob_vs_chi = ranksum(plaques_bob,plaques_chi);
% 
% % vector of p values 
% % p_value_vector = [ MW_bob_vs_chi ; MW_bob_vs_sham ; MW_chi_vs_sham ] ; 
% 
%     filename = 'Left Hemisphere Particle Analysis greater than or equal to 1000um.xlsx';
%     writematrix(plaques_bob',filename,'Sheet','Bobola (m=3)','Range','A2')
%     writematrix(plaques_chi',filename,'Sheet','Chikodi (m=3)','Range','A2')
%     writematrix(plaques_sham',filename,'Sheet','Sham (m=3)','Range','A2')
%     
%     % MWs and summary 
%     writematrix(MW_bob_vs_chi,filename,'Sheet','Summary','Range','B2')
%     writematrix(MW_bob_vs_sham,filename,'Sheet','Summary','Range','C2')
%     writematrix(MW_chi_vs_sham,filename,'Sheet','Summary','Range','D2')
%     writecell({'MW bob vs. chi'},filename,'Sheet','Summary','Range','B1');
%     writecell({'MW bob vs. sham,'},filename,'Sheet','Summary','Range','C1');
%     writecell({'MW chi vs. sham'},filename,'Sheet','Summary','Range','D1');
%     writecell({'p values'},filename,'Sheet','Summary','Range','A2');
%     
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
%    
%     
%     % titling A1 column for each sheet
%     A1title = {'plaque area (um^2)'};
%     writecell(A1title,filename,'Sheet','Bobola (m=3)','Range','A1');
%     writecell(A1title,filename,'Sheet','Chikodi (m=3)','Range','A1');
%     writecell(A1title,filename,'Sheet','Sham (m=3)','Range','A1');
%     writecell(A1title,filename,'Sheet','Sham (m=3)','Range','A1');


end 
