### velocity pptx auto generator
Generate several slides:
1) First slide is simple series chart with data from xlsx/csv file (header: sku,nd,ros)
2) Second slide is the same as first one but with pictures at the same positions inside square. 
Square has border with (x;y)=0;0 at bottom left side and (x;y)=max(nd);max(ros) values at top right corner. 
Because there is no proper method to control scale on chart and legend size.
3) Third one is simple queue of the pictures ordered by ros value in descending order with 15 values in a column 
(if more elements it will continue with 'y'=0 and a new column with shifted 'x' value)
