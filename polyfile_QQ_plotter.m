% Authors: Alissa P., Henry T. 
% this script reads histology data from AD excel spreadsheets and produces QQ plots of plaque sizes for different US protocol cohorts
% it can also create a new consolidated spreadsheet from multiple files 
clc 
close all 
clear all

%% creating new xcel sheet decision
decision = input("Would you like to create a consolidated excel file for the AB plaques by US protocol? ('1'=yes, '0'=no): ") ;

%% reading xcel data 

% bobola 
file1 = xlsread('Bob_11_12_21.tif Plaque Analysis.xlsx');
file2 = xlsread('Bob_12_10_21_LH_ROI plaque analysis.xlsx');
file3 = xlsread('Bob_12_18_LH plaque analysis.xlsx');
bob_plaques = [file1(:,2) ; file2(:,2) ; file3(:,2)];
bob_plaques = (bob_plaques.*1000000)';

% chikodi 
file4 = xlsread('Chi_11_11_21_fullLH-1.tif plaque analysis.xlsx');
file5 = xlsread('Chi_12_03_21_LH_ROI.tif Plaque Analysis.xlsx');
file6 = xlsread('Chi_12_09_21_LH.tif Plaque Analysis.xlsx');
chi_plaques = [file4(:,2) ; file5(:,2) ; file6(:,2)];
chi_plaques = (chi_plaques.*1000000)';

% sham
file7 = xlsread('Sham_M1_AMB_filtered_full_S9.tif Plaque Analysis.xlsx');
file8 = xlsread('Sham_M2_S3_LH_fulltif.tif Plaque Analysis.xlsx');
file9 = xlsread('Sham_M3_LH_ROItif.tif Plaque Analysis.xlsx');
sham_plaques = [file7(:,2) ; file8(:,2) ; file9(:,2)]';
% sham_plaques = (sham_plaques.*1000000)'; % already in um^2 in sheet

%% QQ plots 

% chikodi vs. sham 
figure(1)
chi_vs_sham = qqplot(chi_plaques,sham_plaques) ;
        title('QQ plot of chi vs sham')
        xlabel('Quantiles of Chikodi data')
        ylabel('Quantiles of sham data')
%         hold on 
%         plot(0:100:16000, 0:100:16000) 
% 
% Bobola vs. sham 
figure(2)
bob_vs_sham = qqplot(bob_plaques,sham_plaques) ;
        title('QQ plot of bob vs sham')
        xlabel('Quantiles of bobola data')
        ylabel('Quantiles of sham data')
%         hold on 
%         plot(0:100:16000, 0:100:16000) 
%         
% % Bobola vs. chikodi 
figure(3)
bob_vs_chi = qqplot(bob_plaques,chi_plaques) ;
        title('QQ plot of bob vs chi')
        xlabel('Quantiles of bobola data')
        ylabel('Quantiles of chikodi data')
%         hold on 
%         plot(0:100:16000, 0:100:16000) 
        

%% Mann Whitneys 

MW_chi_vs_sham = ranksum(chi_plaques,sham_plaques);
MW_bob_vs_sham = ranksum(bob_plaques,sham_plaques);
MW_bob_vs_chi = ranksum(bob_plaques,chi_plaques);

%% exporting data to xcel sheet 
if decision == 1 
    filename = 'Left Hemisphere Particle Analysis.xlsx';
    writematrix(bob_plaques',filename,'Sheet','Bobola (m=3)','Range','A2')
    writematrix(chi_plaques',filename,'Sheet','Chikodi (m=3)','Range','A2')
    writematrix(sham_plaques',filename,'Sheet','Sham (m=3)','Range','A2')
    % titling A1 column for each sheet
    A1title = {'plaque area (um^2)'};
    writecell(A1title,filename,'Sheet','Bobola (m=3)','Range','A1');
    writecell(A1title,filename,'Sheet','Chikodi (m=3)','Range','A1');
    writecell(A1title,filename,'Sheet','Sham (m=3)','Range','A1');
end 