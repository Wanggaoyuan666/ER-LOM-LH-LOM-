%% 读取Excel文件中的数据，确定参数的值
% 定义 Excel 文件的文件名 
filename = 'LHLOM.xlsx';  
 
% 读取表格数据
dataOP = readmatrix(filename, 'Sheet', 'Sheet1');
dataPH = readmatrix(filename, 'Sheet', 'Sheet3');

% 读取各个参数： A = kBT/h      B = kBT  f  Ea0  gama = γ    beita = β
A = dataOP(1,2); B = dataOP(2,2); f = dataOP(4,2);
Ea0 = dataOP(5,2); gama = dataOP(6,2); beita = dataOP(7,2); 

% 读取表1的定值pH
pH = dataOP(8,2);
% 读取表3的定值η
overpotential = dataPH(8,2);

% 自由能
G1 = dataOP(9,2);G2_1 = dataOP(10,2);G2_2 = dataOP(11,2);G3_1 = dataOP(12,2);
G3_2 = dataOP(13,2);G4 = dataOP(14,2);G5 = dataOP(15,2);

disp('第1节运行完毕，已获取BV参数！')

%% 对过电势η的规律，进行运算，并填入表格k
% 读取 Excel 文件的第四列（D 列）数据，为过电位η
D_Data = readmatrix(filename,'Sheet',1,'Range', 'D:D');
parameterColumn = D_Data(2:end,1);

% 定义被积函数，a为过电势η，对x即ε在整个能量域作定积分运算
fun1 = @(a) A*exp(-f*(Ea0+gama*G1))*exp((1-beita)*f*(a-0.0592*pH))+...
            A*exp(-f*(Ea0+gama*(G1-0.0592*14)))*exp((1-beita)*f*(a-0.0592*pH))*10^(pH-14);
fun2 = @(a) A*exp(-f*(Ea0-gama*G1))*exp(-beita*f*(a-0.0592*pH))*10^(-pH)+...
            A*exp(-f*(Ea0-gama*(G1-0.0592*14)))*exp(-beita*f*(a-0.0592*pH));
fun3 = @(a) A*exp(-f*(Ea0+gama*G2_1))*exp((1-beita)*f*(a-0.0592*pH))+...
            A*exp(-f*(Ea0+gama*(G2_1-0.0592*14)))*exp((1-beita)*f*(a-0.0592*pH))*10^(pH-14);
fun4 = @(a) A*exp(-f*(Ea0-gama*G2_1))*exp(-beita*f*(a-0.0592*pH))*10^(-pH)+...
            A*exp(-f*(Ea0-gama*(G2_1-0.0592*14)))*exp(-beita*f*(a-0.0592*pH));
fun5 = @(a) A*exp(-f*(Ea0+gama*G2_2))*exp((1-beita)*f*(a-0.0592*pH))+...
            A*exp(-f*(Ea0+gama*(G2_2-0.0592*14)))*exp((1-beita)*f*(a-0.0592*pH))*10^(pH-14);
fun6 = @(a) A*exp(-f*(Ea0-gama*G2_2))*exp(-beita*f*(a-0.0592*pH))*10^(-pH)+...
            A*exp(-f*(Ea0-gama*(G2_2-0.0592*14)))*exp(-beita*f*(a-0.0592*pH));
fun7 = @(a) A*exp(-f*(Ea0+gama*G3_1))*exp((1-beita)*f*(a-0.0592*pH))+...
            A*exp(-f*(Ea0+gama*(G3_1-0.0592*14)))*exp((1-beita)*f*(a-0.0592*pH))*10^(pH-14);
fun8 = @(a) A*exp(-f*(Ea0-gama*G3_1))*exp(-beita*f*(a-0.0592*pH))*10^(-pH)+...
            A*exp(-f*(Ea0-gama*(G3_1-0.0592*14)))*exp(-beita*f*(a-0.0592*pH));
fun9 = @(a) A*exp(-f*(Ea0+gama*G3_2))*exp((1-beita)*f*(a-0.0592*pH))+...
            A*exp(-f*(Ea0+gama*(G3_2-0.0592*14)))*exp((1-beita)*f*(a-0.0592*pH))*10^(pH-14);
fun10 = @(a) A*exp(-f*(Ea0-gama*G3_2))*exp(-beita*f*(a-0.0592*pH))*10^(-pH)+...
             A*exp(-f*(Ea0-gama*(G3_2-0.0592*14)))*exp(-beita*f*(a-0.0592*pH));
fun11 = @(a) A*exp(-f*(Ea0+gama*G4))*exp((1-beita)*f*(a-0.0592*pH))+...
             A*exp(-f*(Ea0+gama*(G4-0.0592*14)))*exp((1-beita)*f*(a-0.0592*pH))*10^(pH-14);
fun12 = @(a) A*exp(-f*(Ea0-gama*G4))*exp(-beita*f*(a-0.0592*pH))*10^(-pH)+...
             A*exp(-f*(Ea0-gama*(G4-0.0592*14)))*exp(-beita*f*(a-0.0592*pH));
fun13 = @(G5) A*exp((-1/B).*(Ea0+gama*G5));
fun14 = @(G5) A*exp((-1/B).*(Ea0-gama*G5));

% 初始化结果列 
resultColumn1 = NaN(length(parameterColumn), 1); 
resultColumn2 = NaN(length(parameterColumn), 1); 
resultColumn3 = NaN(length(parameterColumn), 1); 
resultColumn4 = NaN(length(parameterColumn), 1);
resultColumn5 = NaN(length(parameterColumn), 1); 
resultColumn6 = NaN(length(parameterColumn), 1);
resultColumn7 = NaN(length(parameterColumn), 1); 
resultColumn8 = NaN(length(parameterColumn), 1);
resultColumn9 = NaN(length(parameterColumn), 1); 
resultColumn10 = NaN(length(parameterColumn), 1); 
resultColumn11 = NaN(length(parameterColumn), 1); 
resultColumn12 = NaN(length(parameterColumn), 1);
resultColumn13 = NaN(length(parameterColumn), 1); 
resultColumn14 = NaN(length(parameterColumn), 1);

% 批量计算定积分 
for i = 1:length(parameterColumn) 
    a = parameterColumn(i); % 获取当前参数 
    resultColumn1(i) = fun1(a);
    resultColumn2(i) = fun2(a);
    resultColumn3(i) = fun3(a);
    resultColumn4(i) = fun4(a);
    resultColumn5(i) = fun5(a);
    resultColumn6(i) = fun6(a);
    resultColumn7(i) = fun7(a);
    resultColumn8(i) = fun8(a);
    resultColumn9(i) = fun9(a);
    resultColumn10(i) = fun10(a);
    resultColumn11(i) = fun11(a);
    resultColumn12(i) = fun12(a);
    resultColumn13(i) = fun13(G5);
    resultColumn14(i) = fun14(G5);
end 
 
% 将每步对应的速率写入对应位置 
writematrix(resultColumn1, filename,'Sheet',1, 'Range', 'E2'); 
writematrix(resultColumn2, filename,'Sheet',1, 'Range', 'F2');
writematrix(resultColumn3, filename,'Sheet',1, 'Range', 'G2'); 
writematrix(resultColumn4, filename,'Sheet',1, 'Range', 'H2');
writematrix(resultColumn5, filename,'Sheet',1, 'Range', 'I2'); 
writematrix(resultColumn6, filename,'Sheet',1, 'Range', 'J2');
writematrix(resultColumn7, filename,'Sheet',1, 'Range', 'K2'); 
writematrix(resultColumn8, filename,'Sheet',1, 'Range', 'L2');
writematrix(resultColumn9, filename,'Sheet',1, 'Range', 'M2');
writematrix(resultColumn10, filename,'Sheet',1, 'Range', 'N2');
writematrix(resultColumn11, filename,'Sheet',1, 'Range', 'O2'); 
writematrix(resultColumn12, filename,'Sheet',1, 'Range', 'P2');
writematrix(resultColumn13, filename,'Sheet',1, 'Range', 'Q2'); 
writematrix(resultColumn14, filename,'Sheet',1, 'Range', 'R2');

disp('第2节运行完毕，η对应的k值已计算完毕并填入Sheet1！');

%% 对酸碱度pH的规律，进行运算，并填入表格k
% 读取 Excel 文件的第四列（D 列）数据，为酸碱度pH
D_DatapH = readmatrix(filename,'Sheet',3,'Range', 'D:D');
parameterColumn = D_DatapH(2:end,1);

% 定义被积函数，a为过电势 
fun1 = @(b) A*exp(-f*(Ea0+gama*G1))*exp((1-beita)*f*(overpotential-0.0592*b))+...
            A*exp(-f*(Ea0+gama*(G1-0.0592*14)))*exp((1-beita)*f*(overpotential-0.0592*b))*10^(b-14);
fun2 = @(b) A*exp(-f*(Ea0-gama*G1))*exp(-beita*f*(overpotential-0.0592*b))*10^(-b)+...
            A*exp(-f*(Ea0-gama*(G1-0.0592*14)))*exp(-beita*f*(overpotential-0.0592*b));
fun3 = @(b) A*exp(-f*(Ea0+gama*G2_1))*exp((1-beita)*f*(overpotential-0.0592*b))+...
            A*exp(-f*(Ea0+gama*(G2_1-0.0592*14)))*exp((1-beita)*f*(overpotential-0.0592*b))*10^(b-14);
fun4 = @(b) A*exp(-f*(Ea0-gama*G2_1))*exp(-beita*f*(overpotential-0.0592*b))*10^(-b)+...
            A*exp(-f*(Ea0-gama*(G2_1-0.0592*14)))*exp(-beita*f*(overpotential-0.0592*b));
fun5 = @(b) A*exp(-f*(Ea0+gama*G2_2))*exp((1-beita)*f*(overpotential-0.0592*b))+...
            A*exp(-f*(Ea0+gama*(G2_2-0.0592*14)))*exp((1-beita)*f*(overpotential-0.0592*b))*10^(b-14);
fun6 = @(b) A*exp(-f*(Ea0-gama*G2_2))*exp(-beita*f*(overpotential-0.0592*b))*10^(-b)+...
            A*exp(-f*(Ea0-gama*(G2_2-0.0592*14)))*exp(-beita*f*(overpotential-0.0592*b));
fun7 = @(b) A*exp(-f*(Ea0+gama*G3_1))*exp((1-beita)*f*(overpotential-0.0592*b))+...
            A*exp(-f*(Ea0+gama*(G3_1-0.0592*14)))*exp((1-beita)*f*(overpotential-0.0592*b))*10^(b-14);
fun8 = @(b) A*exp(-f*(Ea0-gama*G3_1))*exp(-beita*f*(overpotential-0.0592*b))*10^(-b)+...
            A*exp(-f*(Ea0-gama*(G3_1-0.0592*14)))*exp(-beita*f*(overpotential-0.0592*b));
fun9 = @(b) A*exp(-f*(Ea0+gama*G3_2))*exp((1-beita)*f*(overpotential-0.0592*b))+...
               A*exp(-f*(Ea0+gama*(G3_2-0.0592*14)))*exp((1-beita)*f*(overpotential-0.0592*b))*10^(b-14);
fun10 = @(b) A*exp(-f*(Ea0-gama*G3_2))*exp(-beita*f*(overpotential-0.0592*b))*10^(-b)+...
             A*exp(-f*(Ea0-gama*(G3_2-0.0592*14)))*exp(-beita*f*(overpotential-0.0592*b));
fun11 = @(b) A*exp(-f*(Ea0+gama*G4))*exp((1-beita)*f*(overpotential-0.0592*b))+...
             A*exp(-f*(Ea0+gama*(G4-0.0592*14)))*exp((1-beita)*f*(overpotential-0.0592*b))*10^(b-14);
fun12 = @(b) A*exp(-f*(Ea0-gama*G4))*exp(-beita*f*(overpotential-0.0592*b))*10^(-b)+...
             A*exp(-f*(Ea0-gama*(G4-0.0592*14)))*exp(-beita*f*(overpotential-0.0592*b));
fun13 = @(G5) A*exp((-1/B).*(Ea0+gama*G5));
fun14 = @(G5) A*exp((-1/B).*(Ea0-gama*G5));

% 初始化结果列 
resultColumn1 = NaN(length(parameterColumn), 1); 
resultColumn2 = NaN(length(parameterColumn), 1); 
resultColumn3 = NaN(length(parameterColumn), 1); 
resultColumn4 = NaN(length(parameterColumn), 1);
resultColumn5 = NaN(length(parameterColumn), 1); 
resultColumn6 = NaN(length(parameterColumn), 1);
resultColumn7 = NaN(length(parameterColumn), 1); 
resultColumn8 = NaN(length(parameterColumn), 1);
resultColumn9 = NaN(length(parameterColumn), 1); 
resultColumn10 = NaN(length(parameterColumn), 1); 
resultColumn11 = NaN(length(parameterColumn), 1); 
resultColumn12 = NaN(length(parameterColumn), 1);
resultColumn13 = zeros(length(parameterColumn), 1); 
resultColumn14 = zeros(length(parameterColumn), 1);

% 批量计算定积分 
for i = 1:length(parameterColumn) 
    b = parameterColumn(i); % 获取当前参数 
    resultColumn1(i) = fun1(b);
    resultColumn2(i) = fun2(b);
    resultColumn3(i) = fun3(b);
    resultColumn4(i) = fun4(b);
    resultColumn5(i) = fun5(b);
    resultColumn6(i) = fun6(b);
    resultColumn7(i) = fun7(b);
    resultColumn8(i) = fun8(b);
    resultColumn9(i) = fun9(b);
    resultColumn10(i) = fun10(b);
    resultColumn11(i) = fun11(b);
    resultColumn12(i) = fun12(b);
    resultColumn13(i) = fun13(G5);
    resultColumn14(i) = fun14(G5);
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
writematrix(resultColumn9, filename,'Sheet',3, 'Range', 'M2');
writematrix(resultColumn10, filename,'Sheet',3, 'Range', 'N2');
writematrix(resultColumn11, filename,'Sheet',3, 'Range', 'O2'); 
writematrix(resultColumn12, filename,'Sheet',3, 'Range', 'P2');
writematrix(resultColumn13, filename,'Sheet',3, 'Range', 'Q2'); 
writematrix(resultColumn14, filename,'Sheet',3, 'Range', 'R2');

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
data_sheet1_M = readmatrix(filename, 'Sheet', 1, 'Range', 'M:M');
data_sheet1_N = readmatrix(filename, 'Sheet', 1, 'Range', 'N:N');
data_sheet1_O = readmatrix(filename, 'Sheet', 1, 'Range', 'O:O');
data_sheet1_P = readmatrix(filename, 'Sheet', 1, 'Range', 'P:P');
data_sheet1_Q = readmatrix(filename, 'Sheet', 1, 'Range', 'Q:Q');
data_sheet1_R = readmatrix(filename, 'Sheet', 1, 'Range', 'R:R');

% 获取数据的长度
num_rows_sheet1 = length(data_sheet1_E);

% 初始化一个足够大的数组来存储要写入表格 2 的数据
data_to_write1 = NaN(6 * num_rows_sheet1, 1);
data_to_write2 = NaN(6 * num_rows_sheet1, 1);
data_to_write3 = NaN(6 * num_rows_sheet1, 1);
data_to_write4 = NaN(6 * num_rows_sheet1, 1);
data_to_write5 = NaN(6 * num_rows_sheet1, 1);
data_to_write6 = NaN(6 * num_rows_sheet1, 1);
data_to_write7 = NaN(6 * num_rows_sheet1, 1);
data_to_write8 = NaN(6 * num_rows_sheet1, 1);
data_to_write9 = NaN(6 * num_rows_sheet1, 1);
data_to_write10 = NaN(6 * num_rows_sheet1, 1);
data_to_write11 = NaN(6 * num_rows_sheet1, 1);
data_to_write12 = NaN(6 * num_rows_sheet1, 1);
data_to_write13 = NaN(6 * num_rows_sheet1, 1);
data_to_write14 = NaN(6 * num_rows_sheet1, 1);

% 按照规则填充数据
for i = 2:num_rows_sheet1
    target_index = 6 * (i - 2) + 1;
    data_to_write1(target_index) = data_sheet1_E(i);
    data_to_write2(target_index) = data_sheet1_F(i);
    data_to_write3(target_index) = data_sheet1_G(i);
    data_to_write4(target_index) = data_sheet1_H(i);
    data_to_write5(target_index) = data_sheet1_I(i);
    data_to_write6(target_index) = data_sheet1_J(i);
    data_to_write7(target_index) = data_sheet1_K(i);
    data_to_write8(target_index) = data_sheet1_L(i);
    data_to_write9(target_index) = data_sheet1_M(i);
    data_to_write10(target_index) = data_sheet1_N(i);
    data_to_write11(target_index) = data_sheet1_O(i);
    data_to_write12(target_index) = data_sheet1_P(i);
    data_to_write13(target_index) = data_sheet1_Q(i);
    data_to_write14(target_index) = data_sheet1_R(i);
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
writematrix(data_to_write9, filename, 'Sheet', 2, 'Range', 'I:I');
writematrix(data_to_write10, filename, 'Sheet', 2, 'Range', 'J:J');
writematrix(data_to_write11, filename, 'Sheet', 2, 'Range', 'K:K');
writematrix(data_to_write12, filename, 'Sheet', 2, 'Range', 'L:L');
writematrix(data_to_write13, filename, 'Sheet', 2, 'Range', 'M:M');
writematrix(data_to_write14, filename, 'Sheet', 2, 'Range', 'N:N');

disp('第3节运行完毕，k值已填入Sheet2！');

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
data_sheet3_M = readmatrix(filename, 'Sheet', 3, 'Range', 'M:M');
data_sheet3_N = readmatrix(filename, 'Sheet', 3, 'Range', 'N:N');
data_sheet3_O = readmatrix(filename, 'Sheet', 3, 'Range', 'O:O');
data_sheet3_P = readmatrix(filename, 'Sheet', 3, 'Range', 'P:P');
data_sheet3_Q = readmatrix(filename, 'Sheet', 3, 'Range', 'Q:Q');
data_sheet3_R = readmatrix(filename, 'Sheet', 3, 'Range', 'R:R');

% 获取数据的长度
num_rows_sheet2 = length(data_sheet3_E);

% 初始化一个足够大的数组来存储要写入表格 2 的数据
data_to_write1 = NaN(6 * num_rows_sheet2, 1);
data_to_write2 = NaN(6 * num_rows_sheet2, 1);
data_to_write3 = NaN(6 * num_rows_sheet2, 1);
data_to_write4 = NaN(6 * num_rows_sheet2, 1);
data_to_write5 = NaN(6 * num_rows_sheet2, 1);
data_to_write6 = NaN(6 * num_rows_sheet2, 1);
data_to_write7 = NaN(6 * num_rows_sheet2, 1);
data_to_write8 = NaN(6 * num_rows_sheet2, 1);
data_to_write9 = NaN(6 * num_rows_sheet2, 1);
data_to_write10 = NaN(6 * num_rows_sheet2, 1);
data_to_write11 = NaN(6 * num_rows_sheet2, 1);
data_to_write12 = NaN(6 * num_rows_sheet2, 1);
data_to_write13 = NaN(6 * num_rows_sheet2, 1);
data_to_write14 = NaN(6 * num_rows_sheet2, 1);

% 按照规则填充数据
for i = 2:num_rows_sheet2
    target_index = 6 * (i - 2) + 1;
    data_to_write1(target_index) = data_sheet3_E(i);
    data_to_write2(target_index) = data_sheet3_F(i);
    data_to_write3(target_index) = data_sheet3_G(i);
    data_to_write4(target_index) = data_sheet3_H(i);
    data_to_write5(target_index) = data_sheet3_I(i);
    data_to_write6(target_index) = data_sheet3_J(i);
    data_to_write7(target_index) = data_sheet3_K(i);
    data_to_write8(target_index) = data_sheet3_L(i);
    data_to_write9(target_index) = data_sheet3_M(i);
    data_to_write10(target_index) = data_sheet3_N(i);
    data_to_write11(target_index) = data_sheet3_O(i);
    data_to_write12(target_index) = data_sheet3_P(i);
    data_to_write13(target_index) = data_sheet3_Q(i);
    data_to_write14(target_index) = data_sheet3_R(i);
end

% 将数据写入表格 2 的 A 列
writematrix(data_to_write1, filename, 'Sheet', 4, 'Range', 'A:A');
writematrix(data_to_write2, filename, 'Sheet', 4, 'Range', 'B:B');
writematrix(data_to_write3, filename, 'Sheet', 4, 'Range', 'C:C');
writematrix(data_to_write4, filename, 'Sheet', 4, 'Range', 'D:D');
writematrix(data_to_write5, filename, 'Sheet', 4, 'Range', 'E:E');
writematrix(data_to_write6, filename, 'Sheet', 4, 'Range', 'F:F');
writematrix(data_to_write7, filename, 'Sheet', 4, 'Range', 'G:G');
writematrix(data_to_write8, filename, 'Sheet', 4, 'Range', 'H:H');
writematrix(data_to_write9, filename, 'Sheet', 4, 'Range', 'I:I');
writematrix(data_to_write10, filename, 'Sheet', 4, 'Range', 'J:J');
writematrix(data_to_write11, filename, 'Sheet', 4, 'Range', 'K:K');
writematrix(data_to_write12, filename, 'Sheet', 4, 'Range', 'L:L');
writematrix(data_to_write13, filename, 'Sheet', 4, 'Range', 'M:M');
writematrix(data_to_write14, filename, 'Sheet', 4, 'Range', 'N:N');

disp('第3节运行完毕，k值已填入Sheet4！')

%% 构建η系数矩阵
filename = 'LHLOM.xlsx'; 
% 使用 readmatrix 读取数据 
data_num = readmatrix(filename,'Sheet' ,2);
 
% 获取数据的行数和列数 
rows = length(data_num)+5;
numRows = rows;

% 初始化一个循环变量，每隔 6 行处理一次 
for i = 1:6:numRows
    if i <= numRows
        data_num(i + 1, 15) = data_num(i, 1);% k1
        data_num(i + 5, 15) = data_num(i, 14);% k-5
        data_num(i + 2, 16) = data_num(i, 3);% k2-1
        data_num(i + 3, 16) = data_num(i, 5);% k2-2
        data_num(i + 1, 17) = data_num(i, 4);% k-2-1
        data_num(i + 4, 17) = data_num(i, 7);% k3-1
        data_num(i + 1, 18) = data_num(i, 6);% k-2-2
        data_num(i + 4, 18) = data_num(i, 9);% k3-2
        data_num(i + 2, 19) = data_num(i, 8);% k-3-1
        data_num(i + 3, 19) = data_num(i, 10);% -k-3-2
        data_num(i + 5, 19) = data_num(i, 11);% k4
        data_num(i + 4, 20) = data_num(i, 12);% k-4
    end
 
    % 计算后写入数据 
    if i <= numRows 
        data_num(i + 1, 16) = -1 * (data_num(i, 2) + data_num(i, 3)+data_num(i,5));
        data_num(i + 2, 17) = -1 * (data_num(i, 4) + data_num(i, 7));
        data_num(i + 3, 18) = -1 * (data_num(i, 6) + data_num(i, 9));
        data_num(i + 4, 19) = -1 * (data_num(i, 8) + data_num(i, 10)+data_num(i,11));
        data_num(i + 5, 20) = -1 * (data_num(i, 12) + data_num(i, 13));
    end
end 
 
% 写入固定值 1 和 0
for i = 2:6:numRows 
    if i <= numRows 
        % 固定值写入 
        data_num(i-1, 15:20) = 1;
        data_num(i+1, 15) = 0;data_num(i+2, 15) = 0;data_num(i+3, 15) = 0;
        data_num(i+3, 16) = 0; data_num(i+4, 16) = 0; 
        data_num(i+2, 17) = 0; data_num(i+4, 17) = 0; 
        data_num(i+1, 18) = 0; data_num(i+4, 18) = 0; 
        data_num(i, 19) = 0; data_num(i+2, 20) = 0; 
        data_num(i, 20) = 0; data_num(i+1, 20) = 0; 
    end 
end  
% 将处理后的数据写回 Excel 文件 
writematrix(data_num,filename ,'Sheet',2);

disp('第4节运行完毕，Sheet2矩阵已构建！');

%% 构建pH系数矩阵
filename = 'LHLOM.xlsx'; 
% 使用 readmatrix 读取数据 
data_num4 = readmatrix(filename,'Sheet',4);

% 获取数据的行数和列数 
rows = length(data_num4)+5;
numRows = rows;

% 初始化一个循环变量，每隔 6 行处理一次 
for i = 1:6:numRows
    if i <= numRows
        data_num4(i + 1, 15) = data_num4(i, 1);% k1
        data_num4(i + 5, 15) = data_num4(i, 14);% k-5
        data_num4(i + 2, 16) = data_num4(i, 3);% k2-1
        data_num4(i + 3, 16) = data_num4(i, 5);% k2-2
        data_num4(i + 1, 17) = data_num4(i, 4);% k-2-1
        data_num4(i + 4, 17) = data_num4(i, 7);% k3-1
        data_num4(i + 1, 18) = data_num4(i, 6);% k-2-2
        data_num4(i + 4, 18) = data_num4(i, 9);% k3-2
        data_num4(i + 2, 19) = data_num4(i, 8);% k-3-1
        data_num4(i + 5, 19) = data_num4(i, 11);% k4
        data_num4(i + 3, 19) = data_num4(i, 10);% -k-3-2
        data_num4(i + 4, 20) = data_num4(i, 12);% k-4
    end
 
    % 计算后写入数据 
    if i <= numRows 
        data_num4(i + 1, 16) = -1 * (data_num4(i, 2) + data_num4(i, 3)+data_num4(i,5));
        data_num4(i + 2, 17) = -1 * (data_num4(i, 4) + data_num4(i, 7));
        data_num4(i + 3, 18) = -1 * (data_num4(i, 6) + data_num4(i, 9));
        data_num4(i + 4, 19) = -1 * (data_num4(i, 8) + data_num4(i, 10)+data_num4(i,11));
        data_num4(i + 5, 20) = -1 * (data_num4(i, 12) + data_num4(i, 13));
    end
end 
 
% 写入固定值 1 和 0
for i = 2:6:numRows 
    if i <= numRows 
        % 固定值写入 
        data_num4(i-1, 15:20) = 1;
        data_num4(i+1, 15) = 0;data_num4(i+2, 15) = 0;data_num4(i+3, 15) = 0;
        data_num4(i+3, 16) = 0; data_num4(i+4, 16) = 0; 
        data_num4(i+2, 17) = 0; data_num4(i+4, 17) = 0; 
        data_num4(i+1, 18) = 0; data_num4(i+4, 18) = 0; 
        data_num4(i, 19) = 0; data_num4(i+2, 20) = 0; 
        data_num4(i, 20) = 0; data_num4(i+1, 20) = 0; 
    end 
end  
% 将处理后的数据写回 Excel 文件 
writematrix(data_num4,filename ,'Sheet',4);

disp('第4节运行完毕，Sheet4矩阵已构建！');

%% 写入系数矩阵对应的参数

% 写入矩阵对应的参数
datacanshu = [ 1 ; 0 ; 0 ; 0 ; 0 ; 0 ]; 

% 写入数据 
writematrix(datacanshu, filename, 'Sheet', 2 , 'Range', 'U1'); 
writematrix(datacanshu, filename, 'Sheet', 4 , 'Range', 'U1'); 

disp('第5节运行完毕，矩阵参数已输入Sheet2！');
disp('第5节运行完毕，矩阵参数已输入Sheet4！');
%% 矩阵计算

% 读取 Sheet2 工作表中系数矩阵 
A_data = readmatrix(filename, 'Sheet', 2 , 'Range', 'O:T'); 
 
% 读取 M1到M6 单元格的数据作为常数项 b 
b_data = readmatrix(filename, 'Sheet', 2 , 'Range', 'U1:U6'); 

% 计算需要求解的方程个数
num_systems = size(A_data,1)/6;

% 预分配结果矩阵
solutions = zeros(num_systems, 6); % 假设解是6*1向量
 
% 遍历每个6*6矩阵 
for k = 1:num_systems
    % 提取当前系数矩阵和常数项 
    start_row = (k-1)*6 + 1;
    end_row = k*6;
    A = A_data(start_row:end_row, :);
    b = b_data;
    
    % 求解并转置结果后储存
    solutions(k, :) = (A \ b)';      % 结果按行存储 
end 
 
% 将结果写入Excel
writematrix(solutions, filename, 'Sheet', 1 , 'Range', 'S2'); 

disp('第6节运行完毕，矩阵已计算完成，覆盖度已填入Sheet1！');


%% 矩阵计算

% 读取 Sheet4 工作表中系数矩阵 
A_data = readmatrix(filename, 'Sheet', 4 , 'Range', 'O:T'); 
 
% 读取 M1到M6 单元格的数据作为常数项 b 
b_data = readmatrix(filename, 'Sheet', 4 , 'Range', 'U1:U6'); 

% 计算需要求解的方程个数
num_systems = size(A_data,1)/6;

% 预分配结果矩阵
solutions = zeros(num_systems, 6); % 假设解是6*1向量
 
% 遍历每个6*6矩阵 
for k = 1:num_systems
    % 提取当前系数矩阵和常数项 
    start_row = (k-1)*6 + 1;
    end_row = k*6;
    A = A_data(start_row:end_row, :);
    b = b_data;
    
    % 求解并转置结果后储存
    solutions(k, :) = (A \ b)';      % 结果按行存储 
end 
 
% 将结果写入Excel
writematrix(solutions, filename, 'Sheet', 3 , 'Range', 'S2'); 

disp('第6节运行完毕，矩阵已计算完成，覆盖度已填入Sheet3！');

%% 计算表观反应速率及log(r)

% 读取整个表格数据 
datar1 = readmatrix(filename, 'Sheet',1); 
 
% 提取某列数据，从第二行开始提取（Excel列号：A = 1）
E = datar1(1:end, 5); F = datar1(1:end, 6); G = datar1(1:end, 7);  
H = datar1(1:end, 8); I = datar1(1:end, 9); J = datar1(1:end, 10);
K = datar1(1:end, 11); L = datar1(1:end, 12); M = datar1(1:end, 13); 
N = datar1(1:end, 14); O = datar1(1:end, 15); P = datar1(1:end, 16); 
Q = datar1(1:end, 17); R = datar1(1:end, 18); S = datar1(1:end, 19);  
T = datar1(1:end, 20); U = datar1(1:end, 21); V = datar1(1:end, 22);  
W = datar1(1:end, 23); X = datar1(1:end, 24); 

% 计算r值
Y_result = E .* S - F .* T; % r1
Z_result = G .* T - H .* U; % r2-1
AA_result = I .* T - J .* V; % r2-2
AB_result = K .* U - L .* W; % r3-1
AC_result = M .* V - N .* W; % r3-2
AD_result = O .* W - P .* X; % r4
AE_result = Q .* X - R .* S; % r5
log1 = log10(abs(Y_result));
log2 = log10(abs(Z_result));
log3 = log10(abs(AA_result));

% 写入r值及log值
writematrix(Y_result,filename, 'Sheet', 1, 'Range','Y2');
writematrix(Z_result,filename, 'Sheet', 1, 'Range','Z2');
writematrix(AA_result,filename, 'Sheet', 1, 'Range','AA2');
writematrix(AB_result,filename, 'Sheet', 1, 'Range','AB2');
writematrix(AC_result,filename, 'Sheet', 1, 'Range','AC2');
writematrix(AD_result,filename, 'Sheet', 1, 'Range','AD2');
writematrix(AE_result,filename, 'Sheet', 1, 'Range','AE2');
writematrix(log1,filename, 'Sheet', 1, 'Range','AF2');
writematrix(log2,filename, 'Sheet', 1, 'Range','AG2');
writematrix(log3,filename, 'Sheet', 1, 'Range','AH2');
disp('第7节运行完毕，log（r）已计算完毕！');

%% 计算表观反应速率及log(r)

% 读取整个表格数据 
datar2 = readmatrix(filename, 'Sheet',3); 
 
% 提取某列数据，从第二行开始提取（Excel列号：A = 1）
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
Q = datar2(1:end, 17); 
R = datar2(1:end, 18);
S = datar2(1:end, 19);  
T = datar2(1:end, 20);  
U = datar2(1:end, 21); 
V = datar2(1:end, 22);  
W = datar2(1:end, 23);  
X = datar2(1:end, 24); 

% 计算r值
Y_result = E .* S - F .* T; % r1
Z_result = G .* T - H .* U; % r2-1
AA_result = I .* T - J .* V; % r2-2
AB_result = K .* U - L .* W; % r3-1
AC_result = M .* V - N .* W; % r3-2
AD_result = O .* W - P .* X; % r4
AE_result = Q .* X - R .* S; % r5
log1 = log10(abs(Y_result));
log2 = log10(abs(Z_result));
log3 = log10(abs(AA_result));

% 写入r值及log值
writematrix(Y_result,filename, 'Sheet', 3, 'Range','Y2');
writematrix(Z_result,filename, 'Sheet', 3, 'Range','Z2');
writematrix(AA_result,filename, 'Sheet', 3, 'Range','AA2');
writematrix(AB_result,filename, 'Sheet', 3, 'Range','AB2');
writematrix(AC_result,filename, 'Sheet', 3, 'Range','AC2');
writematrix(AD_result,filename, 'Sheet', 3, 'Range','AD2');
writematrix(AE_result,filename, 'Sheet', 3, 'Range','AE2');
writematrix(log1,filename, 'Sheet', 3, 'Range','AF2');
writematrix(log2,filename, 'Sheet', 3, 'Range','AG2');
writematrix(log3,filename, 'Sheet', 3, 'Range','AH2');
disp('第7节运行完毕，log（r）已计算完毕！');
