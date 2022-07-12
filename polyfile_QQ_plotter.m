% Authors: Alissa P., Henry T. 
% this script reads histology data from AD excel spreadsheets and produces QQ plots of plaque sizes for different US protocol cohorts
% it can also create a new consolidated spreadsheet from multiple files 
clc 
close all 
clear all

%% creating new xcel sheet decision
decision = input("Would you like to create a consolidated excel file for the AB plaques by US protocol? ('1'=yes, '0'=no): ") ;
decision2 = input("Would you like to create a second sheet of filtered values? ('1'=yes, '0'=no): ") ;
%% reading xcel data 

% bobola 
file1 = xlsread('Data\Bob_11_12_21.tif Plaque Analysis.csv');
file2 = xlsread('Data\Bob_12_10_21_LH_ROI.tif Plaque Analysis.csv');
file3 = xlsread('Data\Bob_12_18_LH.tif Plaque Analysis.csv');
plaques_bob = [file1(:,2) ; file2(:,2) ; file3(:,2)]';
% plaques_bob = (plaques_bob.*1000000)';

% chikodi 
file4 = xlsread('Data\Chi_11_11_21_fullLH-1.tif plaque analysis.csv');
file5 = xlsread('Data\Chi_12_03_21_LH_ROI.tif Plaque Analysis.csv');
file6 = xlsread('Data\Chi_12_09_21_LH.tif Plaque Analysis.csv');
plaques_chi = [file4(:,2) ; file5(:,2) ; file6(:,2)]';
% plaques_chi = (plaques_chi.*1000000)';

% sham
file7 = xlsread('Data\SHAM 12.18.21 ABM S11 full LHf plaque analysis.csv');
file8 = xlsread('Data\Sham_M2_S3_LH_fulltif.tif Plaque Analysis.csv');
file9 = xlsread('Data\Sham_M3_LH_ROItif.tif Plaque Analysis.csv');
plaques_sham = [file7(:,2) ; file8(:,2) ; file9(:,2)]';
% sham_plaques = (sham_plaques.*1000000)'; % already in um^2 in sheet

%% QQ plots 

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
        

%% Mann Whitneys 

MW_chi_vs_sham = ranksum(plaques_chi,plaques_sham);
MW_bob_vs_sham = ranksum(plaques_bob,plaques_sham);
MW_bob_vs_chi = ranksum(plaques_bob,plaques_chi);

median_bob = median(plaques_bob);
median_chi = median(plaques_chi);
median_sham = median(plaques_sham);

mean_bob = mean(plaques_bob);
mean_chi = mean(plaques_chi);
mean_sham = mean(plaques_sham);

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
