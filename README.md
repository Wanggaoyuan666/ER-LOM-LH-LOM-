# ER-LOM-LH-LOM
# 基于ER-LOM和LH-LOM的OER反应的微观动力学模拟计算软件
# 脚本文件分为7个部分与程序界面的回调函数一起实现相关功能——下面以ERLOM-BV脚本解释相关功能
## 脚本第一部分：参数获取
%% 读取Excel文件中的数据，确定参数的值
% 定义 Excel 文件的文件名 
filename = 'ERLOM.xlsx';  
% 读取整个矩阵数据 
dataOP = readmatrix(filename, 'Sheet', 'Sheet1');
dataPH = readmatrix(filename, 'Sheet', 'Sheet3');
% 读取各个参数： A = kBT/h   B = kBT  f=f  Ea0=Ea0  gama = γ  beita = β
A = dataOP(1,2); B = dataOP(2,2); f = dataOP(4,2);
Ea0 = dataOP(5,2); gama = dataOP(6,2); beita = dataOP(7,2); 
% 读取表1的定值pH
pH = dataOP(8,2);
% 读取表3的定值η
overpotential = dataPH(8,2);
% 自由能
G1 = dataOP(9,2);G2 = dataOP(10,2);G3 = dataOP(11,2);G4 = dataOP(12,2);
disp('第1节运行完毕，已获取BV参数！')
