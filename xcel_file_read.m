function [data, sheet_name_array] = xcel_file_read(filename)
% this read an excel file and returns the data for each sheet and an array
% of sheet names 
[~,sheet_name_array]=xlsfinfo(filename);
for k=1:numel(sheet_name_array)
  data{k}=xlsread(filename,sheet_name_array{k});
end

end 