clc
clear
disp('程序开始运行......')
disp("---------------------------------")
%% 程序主要部分
Data = readmatrix("气象数据.xlsx"); % 读取数据
% 创建时间戳
Hour = Data(:,2);
Minute = Data(:,3);
Year = 2024;
Month = 4;
Day = 14;
Timeseries = datetime(Year,Month,Day,Hour,Minute,0);
% 数据读取
Height = Data(:,6); % 海拔
Air_Pressure = Data(:,7); % 气压
Wind_Dir = Data(:,8); % 风向
Wind_Speed = Data(:,9); % 风速
Tem = Data(:,10); % 干球温度
Humidity = Data(:,12); % 相对湿度
Sea_Presure = Data(:,13); % 海平面气压
Time = Data(:,14); % 时间序列
% 拟合海拔与气压、风速、相对湿度、海平面气压、气温的关系
% 海拔与气压
fitResustHA = polyfit(Height,Air_Pressure,1);
YAir = polyval(fitResustHA,Height);
FitinformHA = fitlm(Height,Air_Pressure);
R2HA = num2str(FitinformHA.Rsquared.Ordinary);
figure(1)
plot(Height,Air_Pressure,'o','MarkerFaceColor',[0,0,1])
hold on
plot(Height,YAir,'-','Color',[0,1,0],'LineWidth',1)
title('海拔——气压关系图')
xlabel('海拔（m）'); 
ylabel('气压（hPa）'); 
legend('气压','拟合曲线')
STRDACHA = 'R²=0.99065';
DIMDACHA = [.71 .5 .3 .3];
annotation('textbox',DIMDACHA,'String',STRDACHA,'FitBoxToText','on');
disp("--------------------")
disp("气压——海拔关系图输出成功")
disp(['R²: ',num2str(R2HA)])
% 海拔与风速
fitResustHWS = polyfit(Height,Wind_Speed,1);
YHWS = polyval(fitResustHWS,Height);
FitinformHWS = fitlm(Height,Wind_Speed);
R2HWS = num2str(FitinformHWS.Rsquared.Ordinary);
figure(2)
plot(Height,Wind_Speed,'o','MarkerFaceColor',[0,0,1])
hold on
plot(Height,YHWS,'-','Color',[0,1,0],'LineWidth',1)
title('海拔——风速关系图')
xlabel('海拔（m）'); 
ylabel('风速（m/s）'); 
legend('风速','拟合曲线')
STRDACHWS = 'R²=0.007006';
DIMDACHWS = [.71 .5 .3 .3];
annotation('textbox',DIMDACHWS,'String',STRDACHWS,'FitBoxToText','on');
disp("--------------------")
disp("海拔——风速关系图输出成功")
disp(['R²: ',num2str(R2HWS)])
% 海拔与湿度
fitResustH = polyfit(Height,Humidity,1);
YH = polyval(fitResustH,Height);
FitinformH = fitlm(Height,Humidity);
R2H = num2str(FitinformH.Rsquared.Ordinary);
figure(3)
plot(Height,Humidity,'o','MarkerFaceColor',[0,0,1])
hold on
plot(Height,YH,'-','Color',[0,1,0],'LineWidth',1)
title('海拔——湿度关系图')
xlabel('海拔（m）'); 
ylabel('湿度（%）'); 
legend('湿度','拟合曲线')
STRDACH = 'R²=0.42583';
DIMDACH = [.71 .5 .3 .3];
annotation('textbox',DIMDACH,'String',STRDACH,'FitBoxToText','on');
disp("--------------------")
disp("海拔——湿度关系图输出成功")
disp(['R²: ',num2str(R2H)])
% 海平面气压
fitResustSP = polyfit(Height,Sea_Presure,1);
YSP = polyval(fitResustSP,Height);
FitinformSP = fitlm(Height,Sea_Presure);
R2SP = num2str(FitinformSP.Rsquared.Ordinary);
figure(4)
plot(Height,Sea_Presure,'o','MarkerFaceColor',[0,0,1])
hold on
plot(Height,YSP,'-','Color',[0,1,0],'LineWidth',1)
title('海拔——海平面气压关系图')
xlabel('海拔（m）'); 
ylabel('海平面气压（hPa）'); 
legend('海平面气压','拟合曲线')
STRDACSP = 'R²=0.32947';
DIMDACSP = [.71 .5 .3 .3];
annotation('textbox',DIMDACSP,'String',STRDACSP,'FitBoxToText','on');
disp("--------------------")
disp("海拔——海平面气压关系图输出成功")
disp(['R²: ',num2str(R2SP)])
% 温度
fitResustHT = polyfit(Height,Tem,1);
YHT = polyval(fitResustHT,Height);
FitinformHT = fitlm(Height,Tem);
R2HT = num2str(FitinformHT.Rsquared.Ordinary);
figure(5)
plot(Height,Tem,'o','MarkerFaceColor',[0,0,1])
hold on
plot(Height,YHT,'-','Color',[0,1,0],'LineWidth',1)
title('海拔——气温关系图')
xlabel('海拔（m）'); 
ylabel('气温（℃）'); 
legend('气温','拟合曲线','Location','northwest')
STRDACHT = 'R²=0.53397';
DIMDACHT = [.16 .5 .3 .3];
annotation('textbox',DIMDACHT,'String',STRDACHT,'FitBoxToText','on');
disp("--------------------")
disp("海拔——气温关系图输出成功")
disp(['R²: ',num2str(R2HT)])

% 拟合时间与湿度、温度、风速的关系
% 湿度
fitResustTH = polyfit(Time,Humidity,1);
YTH = polyval(fitResustTH,Time);
FitinformTH = fitlm(Time,Humidity);
R2TH = num2str(FitinformTH.Rsquared.Ordinary);
figure(6)
plot(Timeseries,Humidity,'o','MarkerFaceColor',[0,0,1])
hold on
plot(Timeseries,YTH,'-','Color',[0,1,0],'LineWidth',1)
title('时间——湿度关系图')
xlabel('时间'); 
ylabel('湿度（%）'); 
legend('湿度','拟合曲线')
STRDACTH = 'R²=0.57293';
DIMDACTH = [.71 .5 .3 .3];
annotation('textbox',DIMDACTH,'String',STRDACTH,'FitBoxToText','on');
disp("--------------------")
disp("时间——湿度关系图输出成功")
disp(['R²: ',num2str(R2TH)])
% 温度
fitResustTT = polyfit(Time,Tem,1);
YTT = polyval(fitResustTT,Time);
FitinformTT = fitlm(Time,Tem);
R2TT = num2str(FitinformTT.Rsquared.Ordinary);
figure(7)
plot(Timeseries,Tem,'o','MarkerFaceColor',[0,0,1])
hold on
plot(Timeseries,YTT,'-','Color',[0,1,0],'LineWidth',1)
title('时间——温度关系图')
xlabel('时间'); 
ylabel('温度（℃）'); 
legend('温度','拟合曲线','Location','northwest')
STRDACTT = 'R²=0.88206';
DIMDACTT = [.16 .5 .3 .3];
annotation('textbox',DIMDACTT,'String',STRDACTT,'FitBoxToText','on');
disp("--------------------")
disp("时间——温度关系图输出成功")
disp(['R²: ',num2str(R2TT)])
% 风速
fitResustWS = polyfit(Time,Wind_Speed,1);
YWS = polyval(fitResustWS,Time);
FitinformWS = fitlm(Time,Wind_Speed);
R2WS = num2str(FitinformWS.Rsquared.Ordinary);
figure(8)
plot(Timeseries,Wind_Speed,'o','MarkerFaceColor',[0,0,1])
hold on
plot(Timeseries,YWS,'-','Color',[0,1,0],'LineWidth',1)
title('时间——风速关系图')
xlabel('时间'); 
ylabel('风速（m/s）'); 
legend('风速','拟合曲线')
STRDACWS = 'R²=0.091602';
DIMDACWS = [.71 .5 .3 .3];
annotation('textbox',DIMDACWS,'String',STRDACWS,'FitBoxToText','on');
disp("--------------------")
disp("时间——风速关系图输出成功")
disp(['R²: ',num2str(R2WS)])

% 风玫瑰图
pathFigure= '.\Figures\' ;
figureUnits = 'centimeters';
figureWidth = 18; 
figureHeight = 15;
Options = {'anglenorth',0,...    
           'angleeast',90,...             
           'labels',{'N (0°)','NE (45°)','E (90°)','SE (135°)','S (180°)','SW (225°)','W (270°)','NW (315°)'},... 
           'freqlabelangle',45};
figureHandle = WindRose(Wind_Dir,Wind_Speed,Options);
set(gcf, 'Units', figureUnits, 'Position', [0 0 figureWidth figureHeight]);
set(gca,'FontSize',12,'Fontname', 'Times New Roman');
disp("--------------------")
disp("风玫瑰图图输出成功")
disp("---------------------------------")
disp("程序运行成功")
datetimeE = datetime("now");
disp("程序运行时间：")
disp(datetimeE)






