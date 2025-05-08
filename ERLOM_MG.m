%% 读取Excel文件中的数据，确定参数的值
% 定义 Excel 文件的文件名 
filename = 'ERLOM.xlsx';  
 
% 读取整个矩阵数据 
dataOP = readmatrix(filename, 'Sheet', 'Sheet1');
dataPH = readmatrix(filename, 'Sheet', 'Sheet3');

% 读取各个参数
% A = kBT/h   B = kBT   f=f  Ea0=Ea0  gama = γ  beita = β
A = dataOP(1,2); B = dataOP(2,2);
% 读取表1的定值pH
pH = dataOP(8,2);
% 读取表3的定值η
overpotential = dataPH(8,2);
% 溶剂重组能 C1 = λ1 C2 = λ2 C3 = λ3 C4 = λ4
C1 = dataOP(13,2);C2 = dataOP(14,2);C3 = dataOP(15,2);C4 = dataOP(16,2);
% 自由能
G1 = dataOP(9,2);G2 = dataOP(10,2);G3 = dataOP(11,2);G4 = dataOP(12,2);
% 积分上下限，取正无穷和负无穷
shang = Inf ;xia = -Inf;

disp('第1节运行完毕，已获取Marcus-Gerishcer参数！');
%% 对过电势η的规律，进行运算，并填入表格k1 k-1 k2 k-2 k3 k-3 k4 k-4
 
% 读取 Excel 文件的第四列（D 列）数据，为过电位η
D_Data = readmatrix(filename,'Sheet',1,'Range', 'D:D');
parameterColumn = D_Data(2:end,1);

% 定义被积函数，a为过电势
fun1 = @(x, a) exp(((G1-(a - 0.0592*pH)+x + C1).^2)./(-4*B*C1)).*(1./(1 + exp(-x./B)))+...
               exp(((G1-(a - 0.0592*(pH-14))+x + C1).^2)./(-4*B*C1)).*(1./(1 + exp(-x./B)))*10^(pH-14); 
fun2 = @(x, a) exp(((-G1+(a - 0.0592*pH)-x + C1).^2)./(-4*B*C1)).*(1./(1 + exp(x./B)))*10^(-pH)+...
               exp(((-G1+(a - 0.0592*(pH-14))-x + C1).^2)./(-4*B*C1)).*(1./(1 + exp(x./B)));
fun3 = @(x, a) exp(((G2-(a - 0.0592*pH)+x + C2).^2)./(-4*B*C2)).*(1./(1 + exp(-x./B)))+...
               exp(((G2-(a - 0.0592*(pH-14))+x + C2).^2)./(-4*B*C2)).*(1./(1 + exp(-x./B)))*10^(pH-14); 
fun4 = @(x, a) exp(((-G2+(a - 0.0592*pH)-x + C2).^2)./(-4*B*C2)).*(1./(1 + exp(x./B)))*10^(-pH)+...
               exp(((-G2+(a - 0.0592*(pH-14))-x + C2).^2)./(-4*B*C2)).*(1./(1 + exp(x./B)));
fun5 = @(x, a) exp(((G3-(a - 0.0592*pH)+x + C3).^2)./(-4*B*C3)).*(1./(1 + exp(-x./B)))+...
               exp(((G3-(a - 0.0592*(pH-14))+x + C3).^2)./(-4*B*C3)).*(1./(1 + exp(-x./B)))*10^(pH-14); 
fun6 = @(x, a) exp(((-G3+(a - 0.0592*pH)-x + C3).^2)./(-4*B*C3)).*(1./(1 + exp(x./B)))*10^(-pH)+...
               exp(((-G3+(a - 0.0592*(pH-14))-x + C3).^2)./(-4*B*C3)).*(1./(1 + exp(x./B)));
fun7 = @(x, a) exp(((G4-(a - 0.0592*pH)+x + C4).^2)./(-4*B*C4)).*(1./(1 + exp(-x./B)))+...
               exp(((G4-(a - 0.0592*(pH-14))+x + C4).^2)./(-4*B*C4)).*(1./(1 + exp(-x./B)))*10^(pH-14); 
fun8 = @(x, a) exp(((-G4+(a - 0.0592*pH)-x + C4).^2)./(-4*B*C4)).*(1./(1 + exp(x./B)))*10^(-pH)+...
               exp(((-G4+(a - 0.0592*(pH-14))-x + C4).^2)./(-4*B*C4)).*(1./(1 + exp(x./B)));

% 初始化结果列 
resultColumn1 = zeros(length(parameterColumn), 1); 
resultColumn2 = zeros(length(parameterColumn), 1); 
resultColumn3 = zeros(length(parameterColumn), 1); 
resultColumn4 = zeros(length(parameterColumn), 1);
resultColumn5 = zeros(length(parameterColumn), 1); 
resultColumn6 = zeros(length(parameterColumn), 1);
resultColumn7 = zeros(length(parameterColumn), 1); 
resultColumn8 = zeros(length(parameterColumn), 1);

% 批量计算定积分 
for i = 1:length(parameterColumn) 
    a = parameterColumn(i); % 获取当前参数 
    resultColumn1(i) = A*integral(@(x) fun1(x, a), xia, shang); 
    resultColumn2(i) = A*integral(@(x) fun2(x, a), xia, shang);
    resultColumn3(i) = A*integral(@(x) fun3(x, a), xia, shang); 
    resultColumn4(i) = A*integral(@(x) fun4(x, a), xia, shang);
    resultColumn5(i) = A*integral(@(x) fun5(x, a), xia, shang); 
    resultColumn6(i) = A*integral(@(x) fun6(x, a), xia, shang);
    resultColumn7(i) = A*integral(@(x) fun7(x, a), xia, shang); 
    resultColumn8(i) = A*integral(@(x) fun8(x, a), xia, shang);
end 
 
% 将每步对应的速率写入对应位置 
writematrix(resultColumn1, filename, 'Range', 'E2'); 
writematrix(resultColumn2, filename, 'Range', 'F2');
writematrix(resultColumn3, filename, 'Range', 'G2'); 
writematrix(resultColumn4, filename, 'Range', 'H2');
writematrix(resultColumn5, filename, 'Range', 'I2'); 
writematrix(resultColumn6, filename, 'Range', 'J2');
writematrix(resultColumn7, filename, 'Range', 'K2'); 
writematrix(resultColumn8, filename, 'Range', 'L2');

disp('第2节运行完毕，η对应的k值已计算完毕并填入Sheet1！');

%% 对酸碱度pH在一定范围内进行运算，并填入表格k1 k-1 k2 k-2 k3 k-3 k4 k-4
 
% 读取 Excel 文件的第四列（D 列）数据，为过电位η
D_Data = readmatrix(filename,'Sheet',3 ,'Range', 'D:D');
parameterColumn = D_Data(2:end,1);

% 定义被积函数，a为过电势 
fun1 = @(x, b) exp(((G1-(overpotential - 0.0592*b)+x + C1).^2)./(-4*B*C1)).*(1./(1 + exp(-x./B)))+...
               exp(((G1-(overpotential - 0.0592*(b-14))+x + C1).^2)./(-4*B*C1)).*(1./(1 + exp(-x./B)))*10^(b-14); 
fun2 = @(x, b) exp(((-G1+(overpotential - 0.0592*b)-x + C1).^2)./(-4*B*C1)).*(1./(1 + exp(x./B)))*10^(-b)+...
               exp(((-G1+(overpotential - 0.0592*(b-14))-x + C1).^2)./(-4*B*C1)).*(1./(1 + exp(x./B)));
fun3 = @(x, b) exp(((G2-(overpotential - 0.0592*b)+x + C2).^2)./(-4*B*C2)).*(1./(1 + exp(-x./B)))+...
               exp(((G2-(overpotential - 0.0592*(b-14))+x + C2).^2)./(-4*B*C2)).*(1./(1 + exp(-x./B)))*10^(b-14); 
fun4 = @(x, b) exp(((-G2+(overpotential - 0.0592*b)-x + C2).^2)./(-4*B*C2)).*(1./(1 + exp(x./B)))*10^(-b)+...
               exp(((-G2+(overpotential - 0.0592*(b-14))-x + C2).^2)./(-4*B*C2)).*(1./(1 + exp(x./B)));
fun5 = @(x, b) exp(((G3-(overpotential - 0.0592*b)+x + C3).^2)./(-4*B*C3)).*(1./(1 + exp(-x./B)))+...
               exp(((G3-(overpotential - 0.0592*(b-14))+x + C3).^2)./(-4*B*C3)).*(1./(1 + exp(-x./B)))*10^(b-14); 
fun6 = @(x, b) exp(((-G3+(overpotential - 0.0592*b)-x + C3).^2)./(-4*B*C3)).*(1./(1 + exp(x./B)))*10^(-b)+...
               exp(((-G3+(overpotential - 0.0592*(b-14))-x + C3).^2)./(-4*B*C3)).*(1./(1 + exp(x./B)));
fun7 = @(x, b) exp(((G4-(overpotential - 0.0592*b)+x + C4).^2)./(-4*B*C4)).*(1./(1 + exp(-x./B)))+...
               exp(((G4-(overpotential - 0.0592*(b-14))+x + C4).^2)./(-4*B*C4)).*(1./(1 + exp(-x./B)))*10^(b-14); 
fun8 = @(x, b) exp(((-G4+(overpotential - 0.0592*b)-x + C4).^2)./(-4*B*C4)).*(1./(1 + exp(x./B)))*10^(-b)+...
               exp(((-G4+(overpotential - 0.0592*(b-14))-x + C4).^2)./(-4*B*C4)).*(1./(1 + exp(x./B)));


% 初始化结果列 
resultColumn1 = zeros(length(parameterColumn), 1); 
resultColumn2 = zeros(length(parameterColumn), 1); 
resultColumn3 = zeros(length(parameterColumn), 1); 
resultColumn4 = zeros(length(parameterColumn), 1);
resultColumn5 = zeros(length(parameterColumn), 1); 
resultColumn6 = zeros(length(parameterColumn), 1);
resultColumn7 = zeros(length(parameterColumn), 1); 
resultColumn8 = zeros(length(parameterColumn), 1);

% 批量计算定积分 
for i = 1:length(parameterColumn) 
    b = parameterColumn(i); % 获取当前参数 
    resultColumn1(i) = A*integral(@(x) fun1(x, b), xia, shang); 
    resultColumn2(i) = A*integral(@(x) fun2(x, b), xia, shang);
    resultColumn3(i) = A*integral(@(x) fun3(x, b), xia, shang); 
    resultColumn4(i) = A*integral(@(x) fun4(x, b), xia, shang);
    resultColumn5(i) = A*integral(@(x) fun5(x, b), xia, shang); 
    resultColumn6(i) = A*integral(@(x) fun6(x, b), xia, shang);
    resultColumn7(i) = A*integral(@(x) fun7(x, b), xia, shang); 
    resultColumn8(i) = A*integral(@(x) fun8(x, b), xia, shang);
end 

% 将每步对应的速率写入对应位置 
writematrix(resultColumn1, filename,'Sheet',3, 'Range', 'E2'); 
writematrix(resultColumn2, filename,'Sheet',3, 'Range', 'F2');
writematrix(resultColumn3, filename,'Sheet',3, 'Range', 'G2'); 
writematrix(resultColumn4, filename,'Sheet',3, 'Range', 'H2');
writematrix(resultColumn5, filename,'Sheet',3, 'Range', 'I2'); 
writematrix(resultColumn6, filename,'Sheet',3, 'Range', 'J2');
writematrix(resultColumn7, filename,'Sheet',3, 'Range', 'K2'); 
writematrix(resultColumn8, filename,'Sheet',3, 'Range', 'L2');

disp('第2节运行完毕，pH对应的k值已计算完毕并填入Sheet3！')

%% 将k值填入表格2，方便进行矩阵运算

% 读取表格 1 的 E 列数据
data_sheet1_E = readmatrix(filename, 'Sheet', 1, 'Range', 'E:E');
data_sheet1_F = readmatrix(filename, 'Sheet', 1, 'Range', 'F:F');
data_sheet1_G = readmatrix(filename, 'Sheet', 1, 'Range', 'G:G');
data_sheet1_H = readmatrix(filename, 'Sheet', 1, 'Range', 'H:H');
data_sheet1_I = readmatrix(filename, 'Sheet', 1, 'Range', 'I:I');
data_sheet1_J = readmatrix(filename, 'Sheet', 1, 'Range', 'J:J');
data_sheet1_K = readmatrix(filename, 'Sheet', 1, 'Range', 'K:K');
data_sheet1_L = readmatrix(filename, 'Sheet', 1, 'Range', 'L:L');

% 获取数据的长度
num_rows_sheet1 = length(data_sheet1_E);

% 初始化一个足够大的数组来存储要写入表格 2 的数据
data_to_write1 = NaN(4 * num_rows_sheet1, 1);
data_to_write2 = NaN(4 * num_rows_sheet1, 1);
data_to_write3 = NaN(4 * num_rows_sheet1, 1);
data_to_write4 = NaN(4 * num_rows_sheet1, 1);
data_to_write5 = NaN(4 * num_rows_sheet1, 1);
data_to_write6 = NaN(4 * num_rows_sheet1, 1);
data_to_write7 = NaN(4 * num_rows_sheet1, 1);
data_to_write8 = NaN(4 * num_rows_sheet1, 1);

% 按照规则填充数据
for i = 2:num_rows_sheet1
    target_index = 4 * (i - 2) + 1;
    data_to_write1(target_index) = data_sheet1_E(i);
    data_to_write2(target_index) = data_sheet1_F(i);
    data_to_write3(target_index) = data_sheet1_G(i);
    data_to_write4(target_index) = data_sheet1_H(i);
    data_to_write5(target_index) = data_sheet1_I(i);
    data_to_write6(target_index) = data_sheet1_J(i);
    data_to_write7(target_index) = data_sheet1_K(i);
    data_to_write8(target_index) = data_sheet1_L(i);
end

% 将数据写入表格 2 的 A 列
writematrix(data_to_write1, filename, 'Sheet', 2, 'Range', 'A:A');
writematrix(data_to_write2, filename, 'Sheet', 2, 'Range', 'B:B');
writematrix(data_to_write3, filename, 'Sheet', 2, 'Range', 'C:C');
writematrix(data_to_write4, filename, 'Sheet', 2, 'Range', 'D:D');
writematrix(data_to_write5, filename, 'Sheet', 2, 'Range', 'E:E');
writematrix(data_to_write6, filename, 'Sheet', 2, 'Range', 'F:F');
writematrix(data_to_write7, filename, 'Sheet', 2, 'Range', 'G:G');
writematrix(data_to_write8, filename, 'Sheet', 2, 'Range', 'H:H');
disp('第3节运行完毕，k值已填入Sheet2！')

%% 将k值填入表格4，方便进行矩阵运算

% 读取表格 3 的 E 列数据
data_sheet3_E = readmatrix(filename, 'Sheet', 3, 'Range', 'E:E');
data_sheet3_F = readmatrix(filename, 'Sheet', 3, 'Range', 'F:F');
data_sheet3_G = readmatrix(filename, 'Sheet', 3, 'Range', 'G:G');
data_sheet3_H = readmatrix(filename, 'Sheet', 3, 'Range', 'H:H');
data_sheet3_I = readmatrix(filename, 'Sheet', 3, 'Range', 'I:I');
data_sheet3_J = readmatrix(filename, 'Sheet', 3, 'Range', 'J:J');
data_sheet3_K = readmatrix(filename, 'Sheet', 3, 'Range', 'K:K');
data_sheet3_L = readmatrix(filename, 'Sheet', 3, 'Range', 'L:L');

% 获取数据的长度
num_rows_sheet2 = length(data_sheet3_E);

% 初始化一个足够大的数组来存储要写入表格 2 的数据
datawrite1 = NaN(4 * num_rows_sheet2, 1);
datawrite2 = NaN(4 * num_rows_sheet2, 1);
datawrite3 = NaN(4 * num_rows_sheet2, 1);
datawrite4 = NaN(4 * num_rows_sheet2, 1);
datawrite5 = NaN(4 * num_rows_sheet2, 1);
datawrite6 = NaN(4 * num_rows_sheet2, 1);
datawrite7 = NaN(4 * num_rows_sheet2, 1);
datawrite8 = NaN(4 * num_rows_sheet2, 1);

% 按照规则填充数据
for i = 2:num_rows_sheet2
    target_index = 4 * (i - 2) + 1;
    datawrite1(target_index) = data_sheet3_E(i);
    datawrite2(target_index) = data_sheet3_F(i);
    datawrite3(target_index) = data_sheet3_G(i);
    datawrite4(target_index) = data_sheet3_H(i);
    datawrite5(target_index) = data_sheet3_I(i);
    datawrite6(target_index) = data_sheet3_J(i);
    datawrite7(target_index) = data_sheet3_K(i);
    datawrite8(target_index) = data_sheet3_L(i);
end

% 将数据写入表格 2 的 A 列
writematrix(datawrite1, filename, 'Sheet', 4, 'Range', 'A:A');
writematrix(datawrite2, filename, 'Sheet', 4, 'Range', 'B:B');
writematrix(datawrite3, filename, 'Sheet', 4, 'Range', 'C:C');
writematrix(datawrite4, filename, 'Sheet', 4, 'Range', 'D:D');
writematrix(datawrite5, filename, 'Sheet', 4, 'Range', 'E:E');
writematrix(datawrite6, filename, 'Sheet', 4, 'Range', 'F:F');
writematrix(datawrite7, filename, 'Sheet', 4, 'Range', 'G:G');
writematrix(datawrite8, filename, 'Sheet', 4, 'Range', 'H:H');

disp('第3节运行完毕，k值已填入Sheet4！')
%% 构建系数矩阵
filename = 'ERLOM.xlsx'; 
% 使用 readmatrix 读取数据 
data_num2 = readmatrix(filename,'Sheet' ,2);
 
% 获取数据的行数和列数 
rows = length(data_num2)+3;
numRows = rows;
% 初始化一个循环变量，每隔 4 行处理一次 
for i = 1:4:numRows 
    % 直接复制数据 
    if i  <= numRows 
        data_num2(i + 1, 9) = data_num2(i, 1);
        data_num2(i + 3, 11) = data_num2(i, 5);
        data_num2(i + 1, 11) = data_num2(i, 4); 
        data_num2(i + 2, 10) = data_num2(i, 3);
        data_num2(i + 2, 12) = data_num2(i, 6);
        data_num2(i + 3, 9) = data_num2(i, 8);
    end 
    % 计算后写入数据 
    if i <= numRows 
        data_num2(i + 1, 10) = -1 * (data_num2(i, 2) + data_num2(i, 3));  % B + C 列 -> J 列
        data_num2(i + 2, 11) = -1 * (data_num2(i, 4) + data_num2(i, 5));  % D + E 列 -> K 列
        data_num2(i + 3, 12) = -1 * (data_num2(i, 6) + data_num2(i, 7));  % F + G 列 -> L 列
    end 
 
    % 固定值写入 
    data_num2(i, 9:12) = 1;  % I-L 列写入 1 
end 
 
% 写入特定列的固定值 0 
for i = 2:4:numRows 
    if i <= numRows 
        data_num2(i, 12) = 0;
        data_num2(i+1, 9) = 0;
        data_num2(i+2, 10) = 0;
    end 
end 
% 将处理后的数据写回 Excel 文件 
writematrix(data_num2,filename ,'Sheet','Sheet2');

disp('第4节运行完毕，Sheet2矩阵已构建！');

%% 构建pH系数矩阵
filename = 'ERLOM.xlsx'; 
% 使用 readmatrix 读取数据 
data_num4 = readmatrix(filename,'Sheet',4);

% 获取数据的行数和列数 
rows = length(data_num4);
numRows = rows+3;

% 初始化一个循环变量，每隔 4 行处理一次 
for i = 1:4:numRows 
    % 直接复制数据 
    if i  <= numRows 
        data_num4(i + 1, 9) = data_num4(i, 1);
        data_num4(i + 3, 11) = data_num4(i, 5);
        data_num4(i + 1, 11) = data_num4(i, 4); 
        data_num4(i + 2, 10) = data_num4(i, 3);
        data_num4(i + 2, 12) = data_num4(i, 6);
        data_num4(i + 3, 9) = data_num4(i, 8);
    end 
    % 计算后写入数据 
    if i <= numRows 
        data_num4(i + 1, 10) = -1 * (data_num4(i, 2) + data_num4(i, 3));  % B + C 列 -> J 列
        data_num4(i + 2, 11) = -1 * (data_num4(i, 4) + data_num4(i, 5));  % D + E 列 -> K 列
        data_num4(i + 3, 12) = -1 * (data_num4(i, 6) + data_num4(i, 7));  % F + G 列 -> L 列
    end 
 
    % 固定值写入 
    data_num4(i, 9:12) = 1;  % I-L 列写入 1 
end 
 
% 写入特定列的固定值 0 
for i = 2:4:numRows 
    if i <= numRows 
        data_num4(i, 12) = 0;
        data_num4(i+1, 9) = 0;
        data_num4(i+2, 10) = 0;
    end 
end 
% 将处理后的数据写回 Excel 文件 
writematrix(data_num4,filename ,'Sheet','Sheet4');

disp('第4节运行完毕，Sheet4矩阵已构建！')

%% 写入系数矩阵对应的参数

% 写入矩阵对应的参数
datacanshu = [1; 0; 0 ;0]; 

% 写入数据 
writematrix(datacanshu, filename, 'Sheet', 2 , 'Range', 'M1'); 
writematrix(datacanshu, filename, 'Sheet', 4 , 'Range', 'M1'); 

disp('第5节运行完毕，矩阵参数已输入Sheet2！')
disp('第5节运行完毕，矩阵参数已输入Sheet4！')

%% 矩阵计算

% 读取 Sheet2 工作表中系数矩阵 
A_data = readmatrix(filename, 'Sheet', 2, 'Range', 'I:L'); 
 
% 读取 M1到M4 单元格的数据作为常数项 b 
b_data = readmatrix(filename, 'Sheet', 2, 'Range', 'M1:M4'); 
 
% 计算需要求解的方程组个数 
num_systems = size(A_data, 1)/4;
 
% 预分配结果矩阵（4列 x num_systems行）
solutions = zeros(num_systems, 4);  % 假设解是4x1向量 
 
% 遍历每个4x4矩阵 
for k = 1:num_systems 
    % 提取当前系数矩阵和常数项 
    start_row = (k-1)*4 + 1;
    end_row = k*4;
    A = A_data(start_row:end_row,:);
    b = b_data;
    % 求解并存储（注意转置方式） 
    solutions(k, :) = (A\b)';      % 结果按行存储 
end 
 
% 将结果写入Excel 
writematrix(solutions, filename, 'Sheet', 1, 'Range', 'M2'); 

disp('第6节运行完毕，矩阵已计算完成，覆盖度已填入Sheet1！');

%% 矩阵计算

% 读取 Sheet4 工作表中系数矩阵 
A_data = readmatrix(filename, 'Sheet', 4, 'Range', 'I:L'); 
 
% 读取 M1到M4 单元格的数据作为常数项 b 
b_data = readmatrix(filename, 'Sheet', 4, 'Range', 'M1:M4'); 
 
% 计算需要求解的方程组个数 
num_systems = size(A_data, 1)/4;
 
% 预分配结果矩阵（4列 x num_systems行）
solutions = zeros(num_systems, 4);  % 假设解是4x1向量 
 
% 遍历每个4x4矩阵 
for k = 1:num_systems 
    % 提取当前系数矩阵和常数项 
    start_row = (k-1)*4 + 1;
    end_row = k*4;
    A = A_data(start_row:end_row,:);
    b = b_data;
    % 求解并存储（注意转置方式） 
    solutions(k, :) = (A\b)';      % 结果按行存储 
end 
 
% 将结果写入Excel 
writematrix(solutions, filename, 'Sheet', 3, 'Range', 'M2'); 

disp('第6节运行完毕，矩阵已计算完成，覆盖度已填入Sheet3！');

%% 计算r1、r2、r3、r4和对应的log

% 读取整个表格数据 
datar1 = readmatrix(filename, 'Sheet',1); 
 
% 步骤2：提取E/F/M/N列的数据（Excel列号：E=5, F=6, M=13, N=14）
E = datar1(1:end, 5); 
F = datar1(1:end, 6);  
G = datar1(1:end, 7);  
H = datar1(1:end, 8);
I = datar1(1:end, 9); 
J = datar1(1:end, 10);
K = datar1(1:end, 11);  
L = datar1(1:end, 12);  
M = datar1(1:end, 13); 
N = datar1(1:end, 14);  
O = datar1(1:end, 15);  
P = datar1(1:end, 16); 

Q_result = E .* M - F .* N;
R_result = G .* N - H .* O;
S_result = I .* O - J .* P;
T_result = K .* P - L .* M;
log = log10(abs(Q_result));


% 步骤4：写入Q列（从Q2开始）
writematrix(Q_result,filename, 'Sheet', 1, 'Range','Q2');
writematrix(R_result,filename, 'Sheet', 1, 'Range','R2');
writematrix(S_result,filename, 'Sheet', 1, 'Range','S2');
writematrix(T_result,filename, 'Sheet', 1, 'Range','T2');
writematrix(log ,filename, 'Sheet', 1, 'Range','U2');

disp('第7节运行完毕，log（r）已计算完毕！')

%% 计算r1、r2、r3、r4和对应的log

% 读取整个表格数据 
datar2 = readmatrix(filename, 'Sheet',3); 
 
% 步骤2：提取E/F/M/N列的数据（Excel列号：E=5, F=6, M=13, N=14）
E = datar2(1:end, 5); 
F = datar2(1:end, 6);  
G = datar2(1:end, 7);  
H = datar2(1:end, 8);
I = datar2(1:end, 9); 
J = datar2(1:end, 10);
K = datar2(1:end, 11);  
L = datar2(1:end, 12);  
M = datar2(1:end, 13); 
N = datar2(1:end, 14);  
O = datar2(1:end, 15);  
P = datar2(1:end, 16); 

Q_result = E .* M - F .* N;
R_result = G .* N - H .* O;
S_result = I .* O - J .* P;
T_result = K .* P - L .* M;
log = log10(abs(Q_result));


% 步骤4：写入Q列（从Q2开始）
writematrix(Q_result,filename, 'Sheet', 3, 'Range','Q2');
writematrix(R_result,filename, 'Sheet', 3, 'Range','R2');
writematrix(S_result,filename, 'Sheet', 3, 'Range','S2');
writematrix(T_result,filename, 'Sheet', 3, 'Range','T2');
writematrix(log ,filename, 'Sheet', 3, 'Range','U2');

disp('第7节运行完毕，log（r）已计算完毕！')