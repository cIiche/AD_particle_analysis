% Authors: Alissa P., Henry T. 
% this script reads histology data from AD excel spreadsheets and
% (1) filters based on plaque size (2) produces new xcel sheet by mouse(not cohort)
clc 
close all 
clear all
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
file7 = xlsread('Data\SHAM 12.18.21 ABM S11 full LH plaque analysis.csv');
file8 = xlsread('Data\Sham_M2_S3_LH_fulltif.tif Plaque Analysis.csv');
file9 = xlsread('Data\Sham_M3_LH_ROItif.tif Plaque Analysis.csv');
plaques_sham = [file7(:,2) ; file8(:,2) ; file9(:,2)]';
% sham_plaques = (sham_plaques.*1000000)'; % already in um^2 in sheet

%% filtering plaques to different sizes 

bobm1 = file1(:,2);
bobm2 = file2(:,2); 
bobm3 = file3(:,2); 

chim1 = file4(:,2);
chim2 = file5(:,2); 
chim3 = file6(:,2); 

shamm1 = file7(:,2);
shamm2 = file8(:,2); 
shamm3 = file9(:,2); 
 

% operation = string(input("What is the operation? (ie. 'greater than', 'less than', 'greater than or equal to' etc.): ")) ;
threshold = input("What is the threshold value?: ") ;
% get_filtered_xcel(plaques_bob, plaques_chi, plaques_sham, threshold, operation);

% upper theshold 
% hardcoded for 7_12_22 assignment 
% bobm1 = bobm1(bobm1 <= threshold) ;
% bobm2 = bobm2(bobm2 <= threshold) ;
% bobm3 = bobm3(bobm3 <= threshold) ;
% 
% chim1 = chim1(chim1 <= threshold) ;
% chim2 = chim2(chim2 <= threshold) ;
% chim3 = chim3(chim3 <= threshold) ;
% 
% shamm1 = shamm1(shamm1 <= threshold) ;
% shamm2 = shamm2(shamm2 <= threshold) ;
% shamm3 = shamm3(shamm3 <= threshold) ;
% % lower threshold 
bobm1 = bobm1(bobm1 >= 6000) ;
bobm2 = bobm2(bobm2 >= 6000) ;
bobm3 = bobm3(bobm3 >= 6000) ;

chim1 = chim1(chim1 >= 6000) ;
chim2 = chim2(chim2 >= 6000) ;
chim3 = chim3(chim3 >= 6000) ;

shamm1 = shamm1(shamm1 >= 6000) ;
shamm2 = shamm2(shamm2 >= 6000) ;
shamm3 = shamm3(shamm3 >= 6000) ;

%% producing xcel sheet 

    filename = 'Left Hemisphere Particle Analysis by mouse greater than or equal to ' + string(threshold) + 'um.xlsx';
    writematrix(bobm1,filename,'Sheet','Bobola m1','Range','A2');
    writematrix(bobm2,filename,'Sheet','Bobola m2','Range','A2');
    writematrix(bobm3,filename,'Sheet','Bobola m3','Range','A2');
    
    writematrix(chim1,filename,'Sheet','Chikodi m1','Range','A2');
    writematrix(chim2,filename,'Sheet','Chikodi m2','Range','A2');
    writematrix(chim3,filename,'Sheet','Chikodi m3','Range','A2');
    
    writematrix(shamm1,filename,'Sheet','Sham m1','Range','A2');
    writematrix(shamm2,filename,'Sheet','Sham m2','Range','A2');
    writematrix(shamm3,filename,'Sheet','Sham m3','Range','A2');
    
    % titling A1 column for each sheet
    A1title = {'plaque area (um^2)'};
    writecell(A1title,filename,'Sheet','Bobola m1','Range','A1');
    writecell(A1title,filename,'Sheet','Bobola m2','Range','A1');
    writecell(A1title,filename,'Sheet','Bobola m3','Range','A1');
            
    writecell(A1title,filename,'Sheet','Chikodi m1','Range','A1');
    writecell(A1title,filename,'Sheet','Chikodi m2','Range','A1');
    writecell(A1title,filename,'Sheet','Chikodi m3','Range','A1');
    
    writecell(A1title,filename,'Sheet','Sham m1','Range','A1');
    writecell(A1title,filename,'Sheet','Sham m2','Range','A1');
    writecell(A1title,filename,'Sheet','Sham m3','Range','A1');


