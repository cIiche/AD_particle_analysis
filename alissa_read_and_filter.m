% Authors: Alissa P., Henry T.
% this script reads histology data from AD excel spreadsheets and
% (1) filters based on plaque size (2) produces new xcel sheet with pages for each mouse(not cohort)
% CAN run series_FULL 5 day chronic data plaque sheets 
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
whatdata = input("Run series A ('1') or series B ('2') or series FULL ('3') data?: ");
if whatdata == 1
    seriesdata= 'Series_A' ;
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
    seriesdata= 'Series_B1' ;
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
   
elseif whatdata == 3 % FULL series data
    seriesdata= 'Series_FULL' ;
    file1 = xlsread('Data\Series_FULL\FULL_Chik_M1_AB_ROI_1_ch00.tif AB results table.csv');
    file2 = xlsread('Data\Series_FULL\FULL_Chik_M1_AB_ROI_2_ch00.tif AB results table.csv');
    file3 = xlsread('Data\Series_FULL\FULL_Chik_M2_AB_ROI_1_ch00.tif AB results table.csv');
    file4 = xlsread('Data\Series_FULL\FULL_Chik_M2_AB_ROI_2_ch00.tif AB results table.csv');
    file5 = xlsread('Data\Series_FULL\FULL_Chik_M3_AB_ROI_1_ch00.tif-(Colour_2) AB results table.csv');
    file6 = xlsread('Data\Series_FULL\FULL_Chik_M3_AB_ROI_2_ch00.tif AB results table.csv');
%     plaques_chi = [file1(:,3) ; file2(:,3) ; file3(:,3); file4(:,3); file5(:,3), file6(:,3)]';
   
    file7 = xlsread('Data\Series_FULL\FULL_Sham_M1_AB_ROI_1_ch00.tif AB results table.csv');
    file8 = xlsread('Data\Series_FULL\FULL_Sham_M1_AB_ROI_2_ch00.tif AB results table.csv');
    file9 = xlsread('Data\Series_FULL\FULL_Sham_M2_AB_ROI_1 _ch00.tif AB results table.csv');
    file10 = xlsread('Data\Series_FULL\FULL_Sham_M2_AB_ROI_2_ch00.tif AB results table.csv');
    file11 = xlsread('Data\Series_FULL\FULL_Sham_M3_AB_ROI_1_ch00.tif AB results table.csv');
    file12 = xlsread('Data\Series_FULL\FULL_Sham_M3_AB_ROI_2_ch00.tif AB results table.csv');
%     plaques_sham = [file7(:,3) ; file8(:,3) ; file9(:,3); file10(:,3); file11(:,3), file12(:,3)]';
end

%% boxplots 
% x_label= {'HC';'SZ'};
% toget=[K M];
% grp=[zeros(1,21),ones(1,19)];
% boxplot(toget,grp)
% set(gca,'xticklabel',x_label)

figure(1)
    % ROI1 boxplot 
%     boxplot([file1(:,1), file3(:,1), file5(:,1), file7(:,1), file9(:,1), file11(:,1)],'Notch','on','Labels', {'Chikodi m1 ROI1','Chikodi m2 ROI1', 'Chikodi m3 ROI1','Sham m1 ROI1', 'Sham m2 ROI1', 'Sham m3 ROI1'}) 
%     hold on 
%     title(filename,' ROI1') 
%     ylabel('area (um^2)') 
    
    ROI1_xlab = {'Chikodi m1 ROI1';'Chikodi m2 ROI1'; 'Chikodi m3 ROI1';'Sham m1 ROI1'; 'Sham m2 ROI1'; 'Sham m3 ROI1'};
    ROI1_y = [file1(:,1)' file3(:,1)' file5(:,1)' file7(:,1)' file9(:,1)' file11(:,1)'];
    grp = [zeros(1,length(file1)),ones(1,length(file3)),2*ones(1,length(file5)),3*ones(1,length(file7)),4*ones(1,length(file9)),5*ones(1,length(file11))];
    boxplot(ROI1_y, grp, 'Notch','on') 
    title('ROI1') 
    ylabel('area (um^2)') 
    set(gca,'xticklabel',ROI1_xlab)

figure(2)
    ROI2_xlab = {'Chikodi m1 ROI2';'Chikodi m2 ROI2'; 'Chikodi m3 ROI2';'Sham m1 ROI2'; 'Sham m2 ROI2'; 'Sham m3 ROI2'};
    ROI2_y = [file2(:,1)' file4(:,1)' file6(:,1)' file8(:,1)' file10(:,1)' file12(:,1)'];
    grp = [zeros(1,length(file2)),ones(1,length(file4)),2*ones(1,length(file6)),3*ones(1,length(file8)),4*ones(1,length(file10)),5*ones(1,length(file12))];
    boxplot(ROI2_y, grp, 'Notch','on') 
    title('ROI2') 
    ylabel('area (um^2)') 
    set(gca,'xticklabel',ROI2_xlab) 
    
%% filtering plaques to different sizes
if whatdata == 3
%     mouse_total = input("How many chikodi mice are there?: ") ;
    %%% maybe use this later for different sizes of chorts--the string of
    %%% mice formed can be used for a data structure after filtering
%     for i = 1:mousetotal
%         chi_mice =
%     end

% for different file formats, to do filtering the protocl_mice strings must
% be changed; read_and_filter_flexible is called by whatdata "3" input and
% runs on loops based on the length and plaque entries in the protocol.mX
% variables 
    
    chi.f1 = [file1(:,3) file1(:,9) file1(:,10)];
    chi.f2 = [file2(:,3) file2(:,9) file2(:,10)];
    chi.f3 = [file3(:,3) file3(:,9) file3(:,10)];
    chi.f4 = [file4(:,3) file4(:,9) file4(:,10)];
    chi.f5 = [file5(:,3) file5(:,9) file5(:,10)];
    chi.f6 = [file6(:,3) file6(:,9) file6(:,10)];
    chi_mice = {'chim1' 'chim2' 'chim3' 'chim4' 'chim5' 'chim6'} ;
   
    sham.f1 = [file7(:,3) file7(:,9) file7(:,10)];
    sham.f2 = [file8(:,3) file8(:,9) file8(:,10)];
    sham.f3 = [file9(:,3) file9(:,9) file9(:,10)];
    sham.f4 = [file10(:,3) file10(:,9) file10(:,10)];
    sham.f5 = [file11(:,3) file11(:,9) file11(:,10)];
    sham.f6 = [file12(:,3) file12(:,9) file12(:,10)];
    sham_mice = {'shamm1' 'shamm2' 'shamm3' 'shamm4' 'shamm5' 'shamm6'} ;
else
%     bobm1 = file1(:,2);
%     bobm2 = file2(:,2);
%     bobm3 = file3(:,2);
%     chim1 = file4(:,2);
%     chim2 = file5(:,2);
%     chim3 = file6(:,2);
%     shamm1 = file7(:,2);
%     shamm2 = file8(:,2);
%     shamm3 = file9(:,2);
    bobm1 = [file1(:,2) file1(:,3) file1(:,4)]; % to preserve x and y coordinates of each plaque
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
    if whatdata == 3
        [filtered_chi,filtered_sham] = get_filtered_flexible(chi_mice, sham_mice, chi, sham, x, 'lol', operation);
    else
        [filtered_bobm1,filtered_bobm2,filtered_bobm3,filtered_chim1,filtered_chim2,filtered_chim3,filtered_shamm1,filtered_shamm2,filtered_shamm3] = get_filtered_xcel(bobm1, bobm2, bobm3, chim1, chim2, chim3, shamm1, shamm2, shamm3, x, 'lol', operation);
    end
elseif operation == 5
    x1 = input("What is x1?: ") ;
    x2 = input("What is x2?: ") ;
    if whatdata ==3 
        [filtered_chi,filtered_sham] = get_filtered_flexible(chi_mice, sham_mice, chi, sham, x1, x2, operation);
    else 
        [filtered_bobm1,filtered_bobm2,filtered_bobm3,filtered_chim1,filtered_chim2,filtered_chim3,filtered_shamm1,filtered_shamm2,filtered_shamm3] = get_filtered_xcel(bobm1, bobm2, bobm3, chim1, chim2, chim3, shamm1, shamm2, shamm3, x1, x2, operation);
    end 
end       
    
%% naming xcel sheet, producing sheet(1/2) 
if operation ~= 5
    if operation == 1
        filename = string(seriesdata) + ' LH Particle Analysis greater than ' + string(x) + 'um.xlsx';
    elseif operation == 2
        filename = string(seriesdata) + ' LH Particle Analysis less than ' + string(x) + 'um.xlsx';
    elseif operation ==3
        filename = string(seriesdata) + ' LH Particle Analysis greater than equal to ' + string(x) + 'um.xlsx';
    elseif operation ==4
        filename = string(seriesdata) + ' LH Particle Analysis less than equal to ' + string(x) + 'um.xlsx';
    end
else
    filename = string(seriesdata) + ' LH Particle Analysis by mouse ' + string(x1) + '-' + string(x2) + 'um.xlsx';
end

%% MW for series_FULL data 
if whatdata == 3 
    % merging chi/sham m1-3 ROI1 and ROI2 plaque AREAS into 1 column for MW  
    chi_ROI1 = [filtered_chi.f1(:,1); filtered_chi.f3(:,1); filtered_chi.f5(:,1)];     
    chi_ROI2 = [filtered_chi.f2(:,1); filtered_chi.f4(:,1); filtered_chi.f6(:,1)];  
    
    sham_ROI1 = [filtered_sham.f1(:,1); filtered_sham.f3(:,1); filtered_sham.f5(:,1)];    
    sham_ROI2 = [filtered_sham.f2(:,1); filtered_sham.f4(:,1); filtered_sham.f6(:,1)];  
   
    % making grouped data lengths even 
    % ROI1 
    if length(chi_ROI1) > length(sham_ROI1) % if chi_ROI1 is longer  
        diff_length = length(chi_ROI1) - length(sham_ROI1) ; 
        NaN_vector = NaN(1, diff_length); 
        sham_ROI1 = [sham_ROI1; NaN_vector'];
    else % if sham_ROI1 is longer 
        diff_length = length(sham_ROI1) - length(chi_ROI1); 
        NaN_vector = NaN(1, diff_length); 
        chi_ROI1 = [chi_ROI1; NaN_vector'];
    end 
    % ROI 2    
    if length(chi_ROI2) > length(sham_ROI2) % if chi_ROI1 is longer  
        diff_length = length(chi_ROI2) - length(sham_ROI2) ; 
        NaN_vector = NaN(1, diff_length); 
        sham_ROI2 = [sham_ROI2; NaN_vector'];
    else % if sham_ROI1 is longer 
        diff_length = length(sham_ROI2) - length(chi_ROI2); 
        NaN_vector = NaN(1, diff_length); 
        chi_ROI2 = [chi_ROI2; NaN_vector'];
    end 
    
    % MWs between equal length chi and sham ROI1 and ROI 2 
        MW_ROI1 = ranksum(chi_ROI1, sham_ROI1); 
        MW_ROI2 = ranksum(chi_ROI2, sham_ROI2);
        
 %% boxplots for series_FULL data 
%     figure(1)
%     % ROI1 boxplot 
%     boxplot([chi_ROI1,sham_ROI1],'Notch','on','Labels', {'Chikodi m1-3 ROI1', 'Sham m1-3 ROI1'}) 
%     hold on 
%     title(filename,' ROI1') 
%     ylabel('area (um^2)') 
%     
%     figure(2)
%     % ROI1 boxplot 
%     boxplot([chi_ROI2,sham_ROI2],'Notch','on','Labels', {'Chikodi m1-3 ROI1', 'Sham m1-3 ROI1'}) 
%     hold on 
%     title(filename,' ROI2') 
%     ylabel('area (um^2)') 
    
end % from whatdata==3 

%% writing xcel sheet, producing sheet(2/2) 

if whatdata == 3 % data stored in structure versus variables
   
    writematrix(filtered_chi.f1,filename,'Sheet','Chikodi m1 ROI 1','Range','A2');
    writematrix(filtered_chi.f2,filename,'Sheet','Chikodi m1 ROI 2','Range','A2');
    writematrix(filtered_chi.f3,filename,'Sheet','Chikodi m2 ROI 1','Range','A2');
    writematrix(filtered_chi.f4,filename,'Sheet','Chikodi m2 ROI 2','Range','A2');
    writematrix(filtered_chi.f5,filename,'Sheet','Chikodi m3 ROI 1','Range','A2');
    writematrix(filtered_chi.f6,filename,'Sheet','Chikodi m3 ROI 2','Range','A2');
   
   
    writematrix(filtered_sham.f1,filename,'Sheet','Sham m1 ROI 1','Range','A2');
    writematrix(filtered_sham.f2,filename,'Sheet','Sham m1 ROI 2','Range','A2');
    writematrix(filtered_sham.f3,filename,'Sheet','Sham m2 ROI 1','Range','A2');
    writematrix(filtered_sham.f4,filename,'Sheet','Sham m2 ROI 2','Range','A2');
    writematrix(filtered_sham.f5,filename,'Sheet','Sham m3 ROI 1','Range','A2');
    writematrix(filtered_sham.f6,filename,'Sheet','Sham m3 ROI 2','Range','A2');
   
    % titling A1, B1, C1 column for each sheet
    A1title = {'plaque area (um^2)'};
    B1title = {'X'};
    C1title = {'Y'};
           
%     writecell(A1title,filename,'Sheet','Chikodi m1 ROI 1','Range','A1');
%     writecell(A1title,filename,'Sheet','Chikodi m1 ROI 2','Range','A1');
%     writecell(A1title,filename,'Sheet','Chikodi m2 ROI 1','Range','A1');
%     writecell(A1title,filename,'Sheet','Chikodi m2 ROI 2','Range','A1');
%     writecell(A1title,filename,'Sheet','Chikodi m3 ROI 1','Range','A1');
%     writecell(A1title,filename,'Sheet','Chikodi m3 ROI 2','Range','A1');
%     writecell(A1title,filename,'Sheet','Sham m1 ROI 1','Range','A1');
%     writecell(A1title,filename,'Sheet','Sham m1 ROI 2','Range','A1');
%     writecell(A1title,filename,'Sheet','Sham m2 ROI 1','Range','A1');
%     writecell(A1title,filename,'Sheet','Sham m2 ROI 2','Range','A1');
%     writecell(A1title,filename,'Sheet','Sham m3 ROI 1','Range','A1');
%     writecell(A1title,filename,'Sheet','Sham m3 ROI 2','Range','A1');
    
sheetnames = {'Chikodi m1 ROI 1' 'Chikodi m1 ROI 2' 'Chikodi m2 ROI 1' 'Chikodi m2 ROI 2' 'Chikodi m3 ROI 1' 'Chikodi m3 ROI 2' 'Sham m1 ROI 1' 'Sham m1 ROI 2' 'Sham m2 ROI 1' 'Sham m2 ROI 2' 'Sham m3 ROI 1' 'Sham m3 ROI 2'} ;
    for i = 1:12 
        writecell(A1title,filename,'Sheet',sheetnames{i},'Range','A1');
        writecell(B1title,filename,'Sheet',sheetnames{i},'Range','B1');
        writecell(C1title,filename,'Sheet',sheetnames{i},'Range','C1');
    end 
    
% summary page with MWs 
    writematrix('Summary',filename,'Sheet','Summary','Range','A1');
    writematrix('MW P-value',filename,'Sheet','Summary','Range','B1');
    writematrix('ROI1',filename,'Sheet','Summary','Range','A2');
    writematrix('ROI2',filename,'Sheet','Summary','Range','A3');
    writematrix(MW_ROI1,filename,'Sheet','Summary','Range','B2');
    writematrix(MW_ROI2,filename,'Sheet','Summary','Range','B3');
    
else % other series data 
   
    writematrix(filtered_bobm1,filename,'Sheet','Bobola m1','Range','A2');
    writematrix(filtered_bobm2,filename,'Sheet','Bobola m2','Range','A2');
    writematrix(filtered_bobm3,filename,'Sheet','Bobola m3','Range','A2');
   
    writematrix(filtered_chim1,filename,'Sheet','Chikodi m1','Range','A2');
    writematrix(filtered_chim2,filename,'Sheet','Chikodi m2','Range','A2');
    writematrix(filtered_chim3,filename,'Sheet','Chikodi m3','Range','A2');
   
    writematrix(filtered_shamm1,filename,'Sheet','Sham m1','Range','A2');
    writematrix(filtered_shamm2,filename,'Sheet','Sham m2','Range','A2');
    writematrix(filtered_shamm3,filename,'Sheet','Sham m3','Range','A2');
   
    % titling A1, B1, C1 column for each sheet
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
end
