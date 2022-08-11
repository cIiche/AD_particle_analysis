% Authors: Alissa P., Henry T. 
% this script reads histology data from AD excel spreadsheets and
% (1) filters based on plaque size (2) produces new xcel sheet with pages for each mouse(not cohort)

% CRITERIA 
% - excel sheets read should have plaque sizes in units of qm in column #2
% - arrange 'data' folder in same folder as code, with series A data in
% 'Data' folder, and series B data in 'Series_B1' folder (inside 'Data'
% folder) 
% - filtered excel sheets are produced in same folder as the scripts

clc 
close all 
clear all
%% reading xcel data 

whatdata = input("Run series A ('1') or series B ('2') data?: ") ;
if whatdata == 1 
    series = 'Series_A' ;
    % bobola 
    file1 = xlsread('Data\Bob_11_12_21 ABM Slice 11 LH ROI.tif Plaque Analysis.csv');
    file2 = xlsread('Data\Bob_12_10_21_ABM_S5 LH_ROI.tif Plaque Analysis.csv');
    file3 = xlsread('Data\Bob_12_18_21_ABM_S5 probably LH ROI.tif Plaque Analysis.csv');
    plaques_bob = [file1(:,2) ; file2(:,2) ; file3(:,2)]';
    % plaques_bob = (plaques_bob.*1000000)'; % if units not in um 

    % chikodi 
    file4 = xlsread('Data\Chi_11_11_21_ABM_S5 LH ROI.tif Plaque Analysis.csv');
    file5 = xlsread('Data\Chi_12_03_21 ABM_S5 LH_ROI.tif Plaque Analysis.csv');
    file6 = xlsread('Data\Chi_12_09_21 ABM_S5 LH ROI.tif Plaque Analysis.csv');
    plaques_chi = [file4(:,2) ; file5(:,2) ; file6(:,2)]';
    % plaques_chi = (plaques_chi.*1000000)'; % if units not in um 

    % sham
    file7 = xlsread('Data\SHAM 12.18.21 ABM S11 full LHf.tif (RGB).tif Plaque Analysis.csv');
    file8 = xlsread('Data\Sham_M2_S3A_LH_fulltif.tif Plaque Analysis.csv');
    file9 = xlsread('Data\Sham_M3_S6A LH_ROItif.tif Plaque Analysis.csv');
    plaques_sham = [file7(:,2) ; file8(:,2) ; file9(:,2)]';
elseif whatdata == 2 
    series = 'Series_B1' ;
    % B1 Series data 
    file1 = xlsread('Data\Series_B1\Series_B1 ROI LH Bob 11.12.21 ABM S12 tif.tif Plaque Analysis.csv');
    file2 = xlsread('Data\Series_B1\Series_B1 ROI LH Bob_12.10.21_ABM1_S6 tif.tif Plaque Analysis.csv');
    file3 = xlsread('Data\Series_B1\Series_B1 ROI LH bob_12.18.21_ABM_S6 tif.tif Plaque Analysis.csv');
    plaques_bob = [file1(:,2) ; file2(:,2) ; file3(:,2)]';
    % plaques_bob = (plaques_bob.*1000000)';

    % chikodi 
    file4 = xlsread('Data\Series_B1\Series_B1 ROI LH Chi 11.11.21 ABM_S11.tif Plaque Analysis.csv');
    file5 = xlsread('Data\Series_B1\Series_B1 ROI LH Chi 12.03.21 ABM_S12 tif.tif Plaque Analysis.csv');
    file6 = xlsread('Data\Series_B1\Series_B1 ROI LH Chi 12.09.21 ABM S12 tif.tif Plaque Analysis.csv');
    plaques_chi = [file4(:,2) ; file5(:,2) ; file6(:,2)]';
    % plaques_chi = (plaques_chi.*1000000)';

    % sham
    file7 = xlsread('Data\Series_B1\Series_B1 ROI LH SHAM 12.18.21 ABM S12 tif.tif Plaque Analysis.csv');
    file8 = xlsread('Data\Series_B1\Series_B1 ROI LH Sham m2 ABM S6 tif.tif Plaque Analysis.csv');
    file9 = xlsread('Data\Series_B1\Series_B1 ROI LH sham m3 ABM S3 tif.tif Plaque Analysis.csv');
    plaques_sham = [file7(:,2) ; file8(:,2) ; file9(:,2)]';
end 

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

operation = input("What is the operation?" +newline+"(ie. '1' = '>x', '2' = '<x', '3' = '>=x', '4' = '<=x', '5' = upper and lower x, where 'x' = threshold value(s)): ") ;
if operation ~= 5 
    x = input("What is the threshold value?: ") ;
%     get_filtered_xcel(plaques_bob, plaques_chi, plaques_sham, x, 'lol', operation);
    [filtered_bobm1,filtered_bobm2,filtered_bobm3,filtered_chim1,filtered_chim2,filtered_chim3,filtered_shamm1,filtered_shamm2,filtered_shamm3] = get_filtered_xcel(bobm1, bobm2, bobm3, chim1, chim2, chim3, shamm1, shamm2, shamm3, x, 'lol', operation);
elseif operation == 5 
    x1 = input("What is x1?: ") ;
    x2 = input("What is x2?: ") ; 
    [filtered_bobm1,filtered_bobm2,filtered_bobm3,filtered_chim1,filtered_chim2,filtered_chim3,filtered_shamm1,filtered_shamm2,filtered_shamm3] = get_filtered_xcel(bobm1, bobm2, bobm3, chim1, chim2, chim3, shamm1, shamm2, shamm3, x1, x2, operation);
end 

%% producing xcel sheet 
if operation ~= 5
    if operation == 1 
        filename = string(series) + ' LH Particle Analysis greater than ' + string(x) + 'um.xlsx';
    elseif operation == 2 
        filename = string(series) + ' LH Particle Analysis less than ' + string(x) + 'um.xlsx';
    elseif operation ==3 
        filename = string(series) + ' LH Particle Analysis greater than equal to ' + string(x) + 'um.xlsx';
    elseif operation ==4 
        filename = string(series) + ' LH Particle Analysis less than equal to ' + string(x) + 'um.xlsx';
    end 
else 
    filename = string(series) + ' LH Particle Analysis by mouse ' + string(x1) + '-' + string(x2) + 'um.xlsx';
end 
    writematrix(filtered_bobm1,filename,'Sheet','Bobola m1','Range','A2');
    writematrix(filtered_bobm2,filename,'Sheet','Bobola m2','Range','A2');
    writematrix(filtered_bobm3,filename,'Sheet','Bobola m3','Range','A2');
    
    writematrix(filtered_chim1,filename,'Sheet','Chikodi m1','Range','A2');
    writematrix(filtered_chim2,filename,'Sheet','Chikodi m2','Range','A2');
    writematrix(filtered_chim3,filename,'Sheet','Chikodi m3','Range','A2');
    
    writematrix(filtered_shamm1,filename,'Sheet','Sham m1','Range','A2');
    writematrix(filtered_shamm2,filename,'Sheet','Sham m2','Range','A2');
    writematrix(filtered_shamm3,filename,'Sheet','Sham m3','Range','A2');
    
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


