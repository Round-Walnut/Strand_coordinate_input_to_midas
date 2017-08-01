%This script is used to solve the input of the strand in midas.
clc
clear
filename = 'coordinate.xlsx';
out_filename = 'MCT.txt';

%read file
data = xlsread(filename,'Sheet1');
[data1,textdata] = xlsread(filename,'Sheet1','G:I');


%
id = [];
for i = 1:size(textdata,1);
    a = textdata{i,1};
    if size(a)
        id = [id,i];
    end
end

%写入到命令流文件
fid = fopen(out_filename,'w');
LINE1_1 = 'NAME=';
LINE1_2 = ', ';
LINE1_3 = ', 0, 0, ROUND, 2D';
LINE2 = '        , USER, 0, 0, NO, ';
LINE3 = '        STRAIGHT, 0, 0, 0, X, 0, 0';
LINE4 = '        0, YES, Y, 0';
Y = '        Y=';
Z = '        Z=';
CHAR6 = ', NO, 0,';
CHAR7 = ', NONE, , , , ';


for i = 1:size(id,2)
    first_id = id(i);
    LINE1 = [LINE1_1,textdata{id(i),1},LINE1_2,textdata{id(i),2},LINE1_2,textdata{id(i),3},LINE1_3];
    fprintf(fid,'%s\r\n', LINE1);
    fprintf(fid,'%s\r\n', LINE2);
    fprintf(fid,'%s\r\n', LINE3);
    fprintf(fid,'%s\r\n', LINE4);  
    
    IDY = first_id;
    IDZ = first_id;
    tempy = [];
    while(IDY <= size(data,1)&&data(IDY,4) < 1e10 )
        tempy = [tempy;[data(IDY,4),data(IDY,5)],data(IDY,6)];
        IDY = IDY +1 ;
    end
    if(tempy(1,1) >= tempy(2,1))

        for iii = size(tempy,1):-1:1
            LINE = [Y,num2str(tempy(iii,1)),' ,',num2str(tempy(iii,2)),CHAR6,num2str(tempy(iii,3)),CHAR7];
            fprintf(fid,'%s\r\n', LINE);
        end
    else
        for iii = 1:size(tempy,1)
            LINE = [Y,num2str(tempy(iii,1)),' ,',num2str(tempy(iii,2)),CHAR6,num2str(tempy(iii,3)),CHAR7];
            fprintf(fid,'%s\r\n', LINE);
        end
    end
    tempz = [];
    
    while(IDZ <= size(data,1) && data(IDZ,1) < 1e10)
        tempz = [tempz;[data(IDZ,1),data(IDZ,2)],data(IDZ,3)];
        IDZ = IDZ +1 ;
    end 
    if(tempz(1,1) >= tempz(2,1)) 
       for jjj = size(tempz,1):-1:1
           LINE = [Z,num2str(tempz(jjj,1)),' ,',num2str(tempz(jjj,2)),CHAR6,num2str(tempz(jjj,3)),CHAR7];
           fprintf(fid,'%s\r\n', LINE);   
       end
    else
        for jjj = 1:size(tempz,1)
           LINE = [Z,num2str(tempz(jjj,1)),' ,',num2str(tempz(jjj,2)),CHAR6,num2str(tempz(jjj,3)),CHAR7];
           fprintf(fid,'%s\r\n', LINE);   
       end
    end

end

fclose(fid);
disp('Succsessfully! Find the MCT.txt')



