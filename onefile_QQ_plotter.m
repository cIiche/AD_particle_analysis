% Authors: Alissa P., Henry T. 
% this script reads histology data from AD excel spreadsheets and produces QQ plots of plaque sizes for different US protocol cohorts
clc 
close all 
clear all
%% importing xcel data 

% data1 = xcel_file_read('AB Plaque Quantification (particle analysis) CENTER ROI.xlsx') ;
[data,sheet_names] = xcel_file_read('AB plaque analysis full LH areas.xlsx') ;
% num = xlsread('AB Plaque Quantification (particle analysis) CENTER ROI.xlsx','Bob M1');

% T = readtable('AB Plaque Quantification (particle analysis) CENTER ROI.xlsx');

%% reading xcel sheet plaque sizes - careful; this should change by file
% data{k}(row_index, col_index)

bobola_plaques = data{1}(:,1);

chikodi_plaques = data{2}(:,1);

sham_plaques = data{3}(:,1); 

%% QQ plots 

% chikodi vs. sham 
figure(1)
chi_vs_sham = qqplot(chikodi_plaques,sham_plaques) ;
        title('QQ plot of chi vs sham')
        xlabel('Quantiles of Chikodi data')
        ylabel('Quantiles of sham data')
        hold on 
        plot(0:100:16000, 0:100:16000) 

% Bobola vs. sham 
figure(2)
bob_vs_sham = qqplot(bobola_plaques,sham_plaques) ;
        title('QQ plot of bob vs sham')
        xlabel('Quantiles of bobola data')
        ylabel('Quantiles of sham data')
        hold on 
        plot(0:100:16000, 0:100:16000) 
        
% Bobola vs. chikodi 
figure(3)
bob_vs_chi = qqplot(bobola_plaques,chikodi_plaques) ;
        title('QQ plot of bob vs chi')
        xlabel('Quantiles of bobola data')
        ylabel('Quantiles of chikodi data')
        hold on 
        plot(0:100:16000, 0:100:16000) 
        

%% Mann Whitneys 
chikodi_plaques = chikodi_plaques';
bobola_plaques = bobola_plaques';
sham_plaques = sham_plaques';
MW_chi_vs_sham = ranksum(chikodi_plaques,sham_plaques);
MW_bob_vs_sham = ranksum(bobola_plaques,sham_plaques);
MW_bob_vs_chi = ranksum(bobola_plaques,chikodi_plaques);
a =length(chikodi_plaques)
b =length(bobola_plaques)
c =length(sham_plaques)

% plaques <500um2
chikodi_plaques500 = chikodi_plaques(chikodi_plaques<500);
bobola_plaques500 = bobola_plaques(bobola_plaques<500);
sham_plaques500 = sham_plaques(sham_plaques<500);
MW500_chi_vs_sham = ranksum(chikodi_plaques500,sham_plaques500);
MW500_bob_vs_sham = ranksum(bobola_plaques500,sham_plaques500);
MW500_bob_vs_chi = ranksum(bobola_plaques500,chikodi_plaques500);

% plaques <50
chikodi_plaques50 = chikodi_plaques(chikodi_plaques<50);
bobola_plaques50 = bobola_plaques(bobola_plaques<50);
sham_plaques50 = sham_plaques(sham_plaques<50);
MW50_chi_vs_sham = ranksum(chikodi_plaques50,sham_plaques50);
MW50_bob_vs_sham = ranksum(bobola_plaques50,sham_plaques50);
MW50_bob_vs_chi = ranksum(bobola_plaques50,chikodi_plaques50);

% plaques <100 
chikodi_plaques100 = chikodi_plaques(chikodi_plaques<100);
bobola_plaques100 = bobola_plaques(bobola_plaques<100);
sham_plaques100 = sham_plaques(sham_plaques<100);
MW100_chi_vs_sham = ranksum(chikodi_plaques100,sham_plaques100);
MW100_bob_vs_sham = ranksum(bobola_plaques100,sham_plaques100);
MW100_bob_vs_chi = ranksum(bobola_plaques100,chikodi_plaques100);


