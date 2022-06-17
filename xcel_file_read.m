function [data, sheet_name_array] = xcel_file_read(filename)
% [p_kw,tbl,stats] = kruskalwallis(data,groups);

[~,sheet_name_array]=xlsfinfo(filename);
for k=1:numel(sheet_name_array)
  data{k}=xlsread(filename,sheet_name_array{k});
end

end 