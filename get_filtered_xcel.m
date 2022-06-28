% this function requires input of pre-xcel read bob, chi, and sham plaque vectors 
% limitation: CURRENTLY can only deal with one threshold value -- not both a lower and upper threshold 


function [filtered_plaques_bob, filtered_plaques_chi, filtered_plaques_sham, filtered_median_bob, filtered_median_chi, filtered_median_sham, filtered_mean_bob, filtered_mean_chi, filtered_mean_sham, filtered_MW_bob_vs_sham, filtered_MW_bob_vs_chi, filtered_MW_chi_vs_sham] = get_filtered_xcel(plaques_bob, plaques_chi, plaques_sham, threshold, operation)
if string(operation) == 'less than' 
    filtered_plaques_bob = plaques_bob(plaques_bob < threshold) ;
    filtered_plaques_chi = plaques_chi(plaques_chi < threshold) ;
    filtered_plaques_sham = plaques_sham(plaques_sham < threshold) ;
elseif string(operation) == 'greater than'
    filtered_plaques_bob = plaques_bob(plaques_bob > threshold) ;
    filtered_plaques_chi = plaques_chi(plaques_chi > threshold) ;
    filtered_plaques_sham = plaques_sham(plaques_sham > threshold) ;
elseif string(operation) == 'less than or equal to'
    filtered_plaques_bob = plaques_bob(plaques_bob <= threshold) ;
    filtered_plaques_chi = plaques_chi(plaques_chi <= threshold) ;
    filtered_plaques_sham = plaques_sham(plaques_sham <= threshold) ;
elseif string(operation) == 'greater than or equal to' 
    filtered_plaques_bob = plaques_bob(plaques_bob >= threshold) ;
    filtered_plaques_chi = plaques_chi(plaques_chi >= threshold) ;
    filtered_plaques_sham = plaques_sham(plaques_sham >= threshold) ;
end 

filtered_median_bob = median(filtered_plaques_bob);
filtered_median_chi = median(filtered_plaques_chi);
filtered_median_sham = median(filtered_plaques_sham);

filtered_mean_bob = mean(filtered_plaques_bob);
filtered_mean_chi = mean(filtered_plaques_chi);
filtered_mean_sham = mean(filtered_plaques_sham);

filtered_MW_chi_vs_sham = ranksum(filtered_plaques_chi,filtered_plaques_sham);
filtered_MW_bob_vs_sham = ranksum(filtered_plaques_bob,filtered_plaques_sham);
filtered_MW_bob_vs_chi = ranksum(filtered_plaques_bob,filtered_plaques_chi);

% producing xcel sheet 
%     filename1 = 'Left Hemisphere Particle Analysis ';
%     filename2 = append(filename1, operation); 
    filename = 'Left Hemisphere Particle Analysis ' + string(operation) + ' ' + num2str(threshold) + 'um.xlsx';
    writematrix(filtered_plaques_bob',filename,'Sheet','Bobola (m=3)','Range','A2');
    writematrix(filtered_plaques_chi',filename,'Sheet','Chikodi (m=3)','Range','A2');
    writematrix(filtered_plaques_sham',filename,'Sheet','Sham (m=3)','Range','A2');
    
    % MWs and summary 
    writematrix(filtered_MW_bob_vs_chi,filename,'Sheet','Summary','Range','B2');
    writematrix(filtered_MW_bob_vs_sham,filename,'Sheet','Summary','Range','C2');
    writematrix(filtered_MW_chi_vs_sham,filename,'Sheet','Summary','Range','D2');
    writecell({'MW bob vs. chi'},filename,'Sheet','Summary','Range','B1');
    writecell({'MW bob vs. sham,'},filename,'Sheet','Summary','Range','C1');
    writecell({'MW chi vs. sham'},filename,'Sheet','Summary','Range','D1');
    writecell({'p values'},filename,'Sheet','Summary','Range','A2');
    
    % medians and means 
    writematrix(filtered_median_bob,filename,'Sheet','Summary','Range','B5');
    writematrix(filtered_median_chi,filename,'Sheet','Summary','Range','C5');
    writematrix(filtered_median_sham,filename,'Sheet','Summary','Range','D5');
    
    writematrix(filtered_mean_bob,filename,'Sheet','Summary','Range','B6');
    writematrix(filtered_mean_chi,filename,'Sheet','Summary','Range','C6');
    writematrix(filtered_mean_sham,filename,'Sheet','Summary','Range','D6');
    
    writecell({'bob'},filename,'Sheet','Summary','Range','B4');
    writecell({'chi'},filename,'Sheet','Summary','Range','C4');
    writecell({'sham'},filename,'Sheet','Summary','Range','D4');
    writecell({'medians'},filename,'Sheet','Summary','Range','A5');
    writecell({'means'},filename,'Sheet','Summary','Range','A6');
   
    
    % titling A1 column for each sheet
    A1title = {'plaque area (um^2)'};
    writecell(A1title,filename,'Sheet','Bobola (m=3)','Range','A1');
    writecell(A1title,filename,'Sheet','Chikodi (m=3)','Range','A1');
    writecell(A1title,filename,'Sheet','Sham (m=3)','Range','A1');
    writecell(A1title,filename,'Sheet','Sham (m=3)','Range','A1');


end 