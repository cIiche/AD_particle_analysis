 % Authors: Alissa P., Henry T. 
OUTDATED 11_15_22: can't run SERIES_FULL data and filter it; use alissa_read_and_filter 
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

whatdata = input("Run series A ('1'), series B ('2'), or 'FULL_' data('3')?: ") ;
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
elseif whatdata == 3
    series = 'FULL series' ;
    
    % FULL series chikodi data 
    file1 = xlsread('Data\Series_FULL\FULL_Chik_M1_AB_ROI_1_ch00.tiff AB results table.csv');
    file2 = xlsread('Data\Series_FULL\FULL_Chik_M1_AB_ROI_2_ch00.tiff AB results table.csv');
    file3 = xlsread('Data\Series_FULL\FULL_Chik_M2_AB_ROI_1_ch00.tiff AB results table.csv');
    file4 = xlsread('Data\Series_FULL\FULL_Chik_M2_AB_ROI_2_ch00.tiff AB results table.csv');
    file5 = xlsread('Data\Series_FULL\FULL_Chik_M3_AB_ROI_1_ch00.tiff AB results table.csv');
    file6 = xlsread('Data\Series_FULL\FULL_Chik_M3_AB_ROI_2_ch00.tiff AB results table.csv');
    plaques_chi = [file1(:,3) ; file2(:,3) ; file3(:,3); file4(:,3) ; file5(:,3) ; file6(:,3)]';

    % FULL series SHAM data 
    file7 = xlsread('Data\Series_FULL\FULL_Sham_M1_AB_ROI_1_ch00.tiff AB results table.csv');
    file8 = xlsread('Data\Series_FULL\FULL_Sham_M1_AB_ROI_2_ch00.tiff AB results table.csv');
    file9 = xlsread('Data\Series_FULL\FULL_Sham_M2_AB_ROI_1 _ch00.tiff AB results table.csv');
    file10 = xlsread('Data\Series_FULL\FULL_Sham_M2_AB_ROI_2_ch00.tiff AB results table.csv');
    file11 = xlsread('Data\Series_FULL\FULL_Sham_M3_AB_ROI_1_ch00.tiff AB results table.csv');
    file12 = xlsread('Data\Series_FULL\FULL_Sham_M3_AB_ROI_2_ch00.tiff AB results table.csv');
    plaques_sham = [file7(:,3) ; file8(:,3) ; file9(:,3); file10(:,3) ; file11(:,3) ; file12(:,3)]';
end

%% filtering plaques to different sizes 

% bobm1 = file1(:,2);
% bobm2 = file2(:,2); 
% bobm3 = file3(:,2); 
% 
% chim1 = file4(:,2);
% chim2 = file5(:,2); 
% chim3 = file6(:,2); 
% 
% shamm1 = file7(:,2);
% shamm2 = file8(:,2); 
% shamm3 = file9(:,2); 


if whatdata == 3 
    chim1 = file1(:,3);
    chim2 = file2(:,3); 
    chim3 = file3(:,3);
    chim4 = file4(:,3);
    chim5 = file5(:,3); 
    chim6 = file6(:,3);
    
    shamm1 = file7(:,3) ;
    shamm2 = file8(:,3); 
    shamm3 = file9(:,3) ;
    shamm4 = file7(:,3) ;
    shamm5 = file8(:,3) ; 
    shamm6 = file9(:,3) ;
else % trying to keep x and Y coordinates of plaque; protocolmX = [ plaque, xcoord, ycoord] 
    bobm1 = [file1(:,2) file1(:,3) file1(:,4)];
    bobm2 = [file2(:,2) file2(:,3) file2(:,4)]; 
    bobm3 = [file3(:,2) file3(:,3) file3(:,4)];

    chim1 = [file4(:,2) file4(:,3) file4(:,4)];
    chim2 = [file5(:,2) file5(:,3) file5(:,4)]; 
    chim3 = [file6(:,2) file6(:,3) file6(:,4)]; 

    shamm1 = [file7(:,2) file7(:,3) file7(:,4)];
    shamm2 = [file8(:,2) file8(:,3) file8(:,4)]; 
    shamm3 = [file9(:,2) file9(:,3) file9(:,4)];
end 


operation = input("What is the operation?" +newline+"(ie. '1' = '>x', '2' = '<x', '3' = '>=x', '4' = '<=x', '5' = upper and lower x, where 'x' = threshold value(s)): ") ;
if operation ~= 5 
    x = input("What is the threshold value?: ") ;
%     get_filtered_xcel(plaques_bob, plaques_chi, plaques_sham, x, 'lol', operation);
    if whatdata == 3 % no x and y plaque coordinates 
        [filtered_chim1,filtered_chim2,filtered_chim3,filtered_chim4,filtered_chim5,filtered_chim6,filtered_shamm1,filtered_shamm2,filtered_shamm3,filtered_shamm4,filtered_shamm5,filtered_shamm6] = get_filtered_xcel(bobm1, bobm2, bobm3, chim1, chim2, chim3, shamm1, shamm2, shamm3, x, 'lol', operation);
    else % series A and B
        [filtered_bobm1,filtered_bobm2,filtered_bobm3,filtered_chim1,filtered_chim2,filtered_chim3,filtered_shamm1,filtered_shamm2,filtered_shamm3] = get_filtered_xcel(bobm1, bobm2, bobm3, chim1, chim2, chim3, shamm1, shamm2, shamm3, x, 'lol', operation);
    end 
elseif operation == 5 
    x1 = input("What is x1?: ") ;
    x2 = input("What is x2?: ") ; 
    [filtered_bobm1,filtered_bobm2,filtered_bobm3,filtered_chim1,filtered_chim2,filtered_chim3,filtered_shamm1,filtered_shamm2,filtered_shamm3] = get_filtered_xcel(bobm1, bobm2, bobm3, chim1, chim2, chim3, shamm1, shamm2, shamm3, x1, x2, operation);
end 

%% producing xcel sheet 

% titles 
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
    
    % titling A1, A2, A3 column for each sheet
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

    A2title = {'X'};
    writecell(A2title,filename,'Sheet','Bobola m1','Range','B1');
    writecell(A2title,filename,'Sheet','Bobola m2','Range','B1');
    writecell(A2title,filename,'Sheet','Bobola m3','Range','B1');
            
    writecell(A2title,filename,'Sheet','Chikodi m1','Range','B1');
    writecell(A2title,filename,'Sheet','Chikodi m2','Range','B1');
    writecell(A2title,filename,'Sheet','Chikodi m3','Range','B1');
    
    writecell(A2title,filename,'Sheet','Sham m1','Range','B1');
    writecell(A2title,filename,'Sheet','Sham m2','Range','B1');
    writecell(A2title,filename,'Sheet','Sham m3','Range','B1');

    A3title = {'Y'};
    writecell(A3title,filename,'Sheet','Bobola m1','Range','C1');
    writecell(A3title,filename,'Sheet','Bobola m2','Range','C1');
    writecell(A3title,filename,'Sheet','Bobola m3','Range','C1');
            
    writecell(A3title,filename,'Sheet','Chikodi m1','Range','C1');
    writecell(A3title,filename,'Sheet','Chikodi m2','Range','C1');
    writecell(A3title,filename,'Sheet','Chikodi m3','Range','C1');
    
    writecell(A3title,filename,'Sheet','Sham m1','Range','C1');
    writecell(A3title,filename,'Sheet','Sham m2','Range','C1');
    writecell(A3title,filename,'Sheet','Sham m3','Range','C1');