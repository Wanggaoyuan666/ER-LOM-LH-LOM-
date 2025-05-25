ER-LOM和LH-LOM程序使用说明

一、程序安装

1、上传程序的ERLOM_web.exe和LHLOM_web.exe为不含运行MATLAB Runtime运行环境版本。

2、安装时，双击安装包，按照提示安装，用户可自定义安装位置。

3、若用户电脑未安装运行环境，程序将自动下载对应版本运行环境——MATLAB Runtime2024b版本，用户可自定义MATLAB Runtime2024b运行环境的安装位置，自动下载后继续安装即可。

4、若用户已经安装相关运行环境，则不需要额外下载，安装程序将自动进入下一步安装界面，继续安装完成即可。

5、程序使用：双击快捷方式或用户自定义安装位置的程序打开，在程序左侧参数输入区，输入相关参数，选择相关方法——BV、M、MG，点击相关按钮，即可计算、列表并绘图，在启动程序路径下，会自动生成一个Excel文件，文件中包含相关参数和计算结果。

二、程序文件.mlapp后缀文件说明——以ERLOM.mlapp为例说明

1、程序文件为MATLAB Appdesigner专属文件，需安装MATLAB软件，启动后双击文件，即可打开。

2、文件中包含两部分内容，分别为设计视图和代码视图。

3、设计视图为程序的界面，可自由添加组件，调整程序界面。

4、代码视图为组件对应的代码，要实现相关功能，需要对对应组件添加对应的回调函数。

三、回调函数相关代码说明——以MATLAB打开ERLOM.mlapp文件后，在代码视图可见相关代码

1、构建Excel文件，填入相关的标题行，规定参数、自变量、计算结果的位置，方便后期执行脚本计算后，相关数据的填入。

    %% 创建ERLOM数据输出Excel表格文件
    filename = 'ERLOM.xlsx';  % 指定要删除的文件名 
    if exist("ERLOM.xlsx", 'file') == 2 % 检查文件是否存在 
        delete(filename); % 删除文件 
    end 
    % 创建空白表格
    file = 'ERLOM.xlsx'; 
    A = []; B = []; C =[]; D = [];
    writematrix(A,file,"Sheet",'Sheet1');
    writematrix(B,file,"Sheet",'Sheet2');
    writematrix(C,file,"Sheet",'Sheet3');
    writematrix(D,file,"Sheet",'Sheet4');
    % 标题行和相关参数
    TitleRow1 = {'参数','数值','单位','η','k1','k-1','k2','k-2','k3','k-3','k4','k-4',...
                'Θ-OMV','Θ-OMOH','ΘOMO','ΘOMOOH','r1','r2','r3','r4','log(r)'};
    TitleRow2 = {'参数','数值','单位','pH','k1','k-1','k2','k-2','k3','k-3','k4','k-4',...
                'Θ-OMV','Θ-OMOH','ΘOMO','ΘOMOOH','r1','r2','r3','r4','log(r)'};
    CanShuNameOP = {'kBT/h','kBT','T','f','Ea0','γ','β','pH','ΔG1','ΔG2','ΔG3','ΔG4',...
                'λ1','λ2','λ3','λ4'}';
    CanShuNamepH = {'kBT/h','kBT','T','f','Ea0','γ','β','η','ΔG1','ΔG2','ΔG3','ΔG4',...
                'λ1','λ2','λ3','λ4'}';
    CanShuData = [6220889369003.16;0.0256925694830201;298.15;38.9217590969594];
    CanShuUnit = {'s-1','eV','K','V-1','eV','','','','eV','eV','eV','eV','eV','eV','eV','eV'}';
    % 写入随η变化表格参数
    writecell(TitleRow1,file,"Sheet",1,'Range','A1');
    writecell(CanShuNameOP,file,"Sheet",1,'Range','A2');
    writematrix(CanShuData,file,"Sheet",1,'Range','B2');
    writecell(CanShuUnit,file,"Sheet",1,'Range','C2');
    % 写入随pH变化表格参数
    writecell(TitleRow2,file,"Sheet",3,'Range','A1');
    writecell(CanShuNamepH,file,"Sheet",3,'Range','A2');
    writematrix(CanShuData,file,"Sheet",3,'Range','B2');
    writecell(CanShuUnit,file,"Sheet",3,'Range','C2');

2、获取用户输入的参数值，并将相关参数填入Excel文件中的对应位置，方便执行脚本时，脚本可自动获取对应单元格的参数。

    %% 获取参数，填入表格对应位置
    filename = 'ERLOM.xlsx';
    yita = (app.OPdown.Value:app.OPstep.Value:app.OPup.Value)';
    pH = (app.pHdown.Value:app.pHstep.Value:app.pHup.Value)';
    writematrix(app.G1EditField.Value,filename,'Sheet',1, 'Range', 'B10');
    writematrix(app.G2EditField.Value,filename,'Sheet',1, 'Range', 'B11');
    writematrix(app.G3EditField.Value,filename,'Sheet',1, 'Range', 'B12');
    writematrix(app.G4EditField.Value,filename,'Sheet',1, 'Range', 'B13');
    writematrix(app.G1EditField.Value,filename,'Sheet',3, 'Range', 'B10');
    writematrix(app.G2EditField.Value,filename,'Sheet',3, 'Range', 'B11');
    writematrix(app.G3EditField.Value,filename,'Sheet',3, 'Range', 'B12');
    writematrix(app.G4EditField.Value,filename,'Sheet',3, 'Range', 'B13');
    writematrix(yita,filename,'Sheet', 1, 'Range', 'D2');
    writematrix(pH,filename,'Sheet', 3, 'Range', 'D2');
    writematrix(app.pHding.Value,filename,'Sheet',1, 'Range', 'B9');
    writematrix(app.yitading.Value,filename,'Sheet',3, 'Range', 'B9');
    writematrix(app.Ea0.Value,filename,'Sheet', 1 , 'Range', 'B6');
    writematrix(app.gama.Value,filename,'Sheet', 1 , 'Range', 'B7');
    writematrix(app.beita.Value,filename,'Sheet', 1 , 'Range', 'B8');
    writematrix(app.Ea0.Value,filename,'Sheet', 3 , 'Range', 'B6');
    writematrix(app.gama.Value,filename,'Sheet', 3 , 'Range', 'B7');
    writematrix(app.beita.Value,filename,'Sheet', 3 , 'Range', 'B8');

3、运行脚本，脚本运行成功后，自动获取Excel表格中计算后的数据，并将数据填入UI界面的表格组件中，并让图组件自动获取相关数据，剔除不合理数据并自动绘图。

    run("ERLOM_BV.m");
    t1 = readtable("ERLOM.xlsx",'VariableNamingRule', 'preserve',"Sheet",1,"Range",'D:U');
    t2 = readtable("ERLOM.xlsx",'VariableNamingRule', 'preserve',"Sheet",3,"Range",'D:U');
    app.tableOP.Data = t1;
    app.tablepH.Data = t2
    % 获取η数据作为x轴
    x1 = table2array(t1(:,1));
    % 获取log(r)作为y轴
    y1 = table2array(t1(:,18));
    % 获取pH数据作为x轴
    x2 = table2array(t2(:,1));
    % 获取log(r)作为y轴
    y2 = table2array(t2(:,18));
    % 获取覆盖度数值作为y轴
    y3 = table2array(t1(:,10));y4 = table2array(t1(:,11));y5 = table2array(t1(:,12));
    y6 = table2array(t1(:,13));y7 = table2array(t2(:,10));y8 = table2array(t2(:,11));
    y9 = table2array(t2(:,12));y10 = table2array(t2(:,13));
    % 清除异常数据 
    y1max = 200;y2max = 200;
    valid_indices1 = y1 < y1max; valid_indices2 = y2 < y2max; % 生成逻辑索引 
    x1_clean = x1(valid_indices1);
    y1_clean = y1(valid_indices1);
    x2_clean = x2(valid_indices2);
    y2_clean = y2(valid_indices2);
    % 清除坐标轴并绘制过电势η图形 
    cla(app.picOP);
    plot(app.picOP, x1_clean, y1_clean);
    cla(app.picOPc);
    plot(app.picOPc,x1,y3,'r',x1,y4,'b',x1,y5,'g',x1,y6,'y');
    % 清除坐标轴并绘制酸碱度pH图形 
    cla(app.picpH);
    plot(app.picpH, x2_clean, y2_clean);
    cla(app.picpHc);
    plot(app.picpHc,x2,y7,'r',x2,y8,'b',x2,y9,'g',x2,y10,'y');
    % 添加图例 
    legend(app.picOPc,'θOMV', 'θOMOH', 'θOMO', 'θOMOOH');
    legend(app.picpHc,'θOMV', 'θOMOH', 'θOMO', 'θOMOOH');
    % 添加坐标轴标签 
    app.picOP.XLabel.String = 'η';app.picOP.YLabel.String = 'log(r)';app.picOP.Title.String = 'η-log(r)';
    app.picOPc.XLabel.String = 'η';app.picOPc.YLabel.String = 'θ';app.picOPc.Title.String = 'η-θ';
    app.picpH.XLabel.String = 'pH';app.picpH.YLabel.String = 'log(r)';app.picpH.Title.String = 'pH-log(r)';
    app.picpHc.XLabel.String = 'pH';app.picpHc.YLabel.String = 'θ';app.picpHc.Title.String = 'pH-θ';

四、脚本文件.m后缀文件说明

脚本文件对应程序中的相关按钮，例如 ERLOM_BV.m 脚本文件对应 ERLOM.exe 中的BV按钮，点击相关按钮后，程序自动运行相关脚本，实现脚本功能，因此共有脚本.m文件六个。

以下以 ERLOM_BV.m 为例，仅以固定酸碱度pH，过电位η为自变量的相关代码进行展示说明，固定过电位η，酸碱度pH为自变量的相关代码重复，不再展示。

1、第一节——获取程序创建表格中的相关参数

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

2、第二节——定义各个基元反应的正逆反应的速率常数表达式，对于ERLOM，有四个基元反应，故定义八个表达式并计算

    %% 对过电势η的规律，进行运算，并填入表格k1 k-1 k2 k-2 k3 k-3 k4 k-4
    % 读取 Excel 文件的第四列（D 列）数据，为过电位η
    D_Data = readmatrix(filename,'Sheet', 1 ,'Range', 'D:D');
    parameterColumn = D_Data(2:end,1);
    % 定义BV动力学的速率常数表达式
    fun1 = @(a) A*exp(-f*(Ea0+gama*G1))*exp((1-beita)*f*(a-0.0592*pH))+...
                A*exp(-f*(Ea0+gama*(G1-0.0592*14)))*exp((1-beita)*f*(a-0.0592*pH))*10^(pH-14);
    fun2 = @(a) A*exp(-f*(Ea0-gama*G1))*exp(-beita*f*(a-0.0592*pH))*10^(-pH)+...
                A*exp(-f*(Ea0-gama*(G1-0.0592*14)))*exp(-beita*f*(a-0.0592*pH));
    fun3 = @(a) A*exp(-f*(Ea0+gama*G2))*exp((1-beita)*f*(a-0.0592*pH))+...
                A*exp(-f*(Ea0+gama*(G2-0.0592*14)))*exp((1-beita)*f*(a-0.0592*pH))*10^(pH-14);
    fun4 = @(a) A*exp(-f*(Ea0-gama*G2))*exp(-beita*f*(a-0.0592*pH))*10^(-pH)+...
                A*exp(-f*(Ea0-gama*(G2-0.0592*14)))*exp(-beita*f*(a-0.0592*pH));
    fun5 = @(a) A*exp(-f*(Ea0+gama*G3))*exp((1-beita)*f*(a-0.0592*pH))+...
                A*exp(-f*(Ea0+gama*(G3-0.0592*14)))*exp((1-beita)*f*(a-0.0592*pH))*10^(pH-14);
    fun6 = @(a) A*exp(-f*(Ea0-gama*G3))*exp(-beita*f*(a-0.0592*pH))*10^(-pH)+...
                A*exp(-f*(Ea0-gama*(G3-0.0592*14)))*exp(-beita*f*(a-0.0592*pH));
    fun7 = @(a) A*exp(-f*(Ea0+gama*G4))*exp((1-beita)*f*(a-0.0592*pH))+...
                A*exp(-f*(Ea0+gama*(G4-0.0592*14)))*exp((1-beita)*f*(a-0.0592*pH))*10^(pH-14);
    fun8 = @(a) A*exp(-f*(Ea0-gama*G4))*exp(-beita*f*(a-0.0592*pH))*10^(-pH)+...
                A*exp(-f*(Ea0-gama*(G4-0.0592*14)))*exp(-beita*f*(a-0.0592*pH));
    % 初始化结果列——方便储存批量计算的各个过电势条件下的速率常数
    resultColumn1 = zeros(length(parameterColumn), 1); 
    resultColumn2 = zeros(length(parameterColumn), 1); 
    resultColumn3 = zeros(length(parameterColumn), 1); 
    resultColumn4 = zeros(length(parameterColumn), 1);
    resultColumn5 = zeros(length(parameterColumn), 1); 
    resultColumn6 = zeros(length(parameterColumn), 1);
    resultColumn7 = zeros(length(parameterColumn), 1); 
    resultColumn8 = zeros(length(parameterColumn), 1);
    % 批量计算k值——使用for循环计算
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
    end 
    % 将每步对应的速率写入Excel表格中对应位置
    writematrix(resultColumn1, filename,'Sheet',1, 'Range', 'E2'); 
    writematrix(resultColumn2, filename,'Sheet',1, 'Range', 'F2');
    writematrix(resultColumn3, filename,'Sheet',1, 'Range', 'G2'); 
    writematrix(resultColumn4, filename,'Sheet',1, 'Range', 'H2');
    writematrix(resultColumn5, filename,'Sheet',1, 'Range', 'I2'); 
    writematrix(resultColumn6, filename,'Sheet',1, 'Range', 'J2');
    writematrix(resultColumn7, filename,'Sheet',1, 'Range', 'K2'); 
    writematrix(resultColumn8, filename,'Sheet',1, 'Range', 'L2');
    disp('第2节运行完毕，η对应的k值已计算完毕并填入Sheet1！');

3、第三节——将计算所得不同条件下的k值填入Excel表格中的新的Sheet中，方便后续批量构建矩阵

    %% 将不同η条件下的k值填入表格2，方便进行矩阵运算
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

4、第四节——将第三步复制的k值批量填入对应矩阵的相关位置，构建多个4*4矩阵

    %% 构建η系数矩阵
    filename = 'ERLOM.xlsx'; 
    % 使用 readmatrix 读取η和pH数据 
    data_num2 = readmatrix(filename,'Sheet',2);
    % 获取数据的行数和列数 
    rows = length(data_num2);
    numRows = rows+3;
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
    writematrix(data_num2,filename ,'Sheet',2);
    disp('第4节运行完毕，Sheet2矩阵已构建！');

5、第五节——写入矩阵对应的参数

    %% 写入系数矩阵对应的参数
    % 写入矩阵对应的参数
    datacanshu = [1; 0; 0 ;0]; 
    % 写入数据 
    writematrix(datacanshu, filename, 'Sheet', 2 , 'Range', 'M1'); 
    writematrix(datacanshu, filename, 'Sheet', 4 , 'Range', 'M1'); 
    disp('第5节运行完毕，矩阵参数已输入Sheet2！')
    disp('第5节运行完毕，矩阵参数已输入Sheet4！')

6、第六节——使用for循环，批量求解每个η条件下的矩阵，求解四种表面活性物种的覆盖度

    %% 矩阵计算
    % 读取 Sheet2 工作表中系数矩阵 
    A_data = readmatrix(filename, 'Sheet', 2, 'Range', 'I:L'); 
    % 读取 M1到M4 单元格的数据作为常数项 b 
    b_data = readmatrix(filename, 'Sheet', 2, 'Range', 'M1:M4'); 
    % 计算需要求解的方程组个数 
    num_systems = size(A_data, 1)/4;
    % 预分配结果矩阵（4列 x num_systems行）
    solutions = zeros(num_systems, 4);  % 假设解是4x4向量 
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

7、第七节——计算各基元反应的表观速率r及其log（r）

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
