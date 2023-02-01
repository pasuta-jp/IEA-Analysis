%変数クリア
clear all
%箱ひげ図と散布図の処理を飛ばすか決定
spider_quick=input('レーダーチャートだけ出したいですか？(Yes→１, No→2)：');

%データ読み込み
DATA=readcell('C:\Users\Akama\Desktop\works\internship\JICA\IEAデータベース\編集マクロ\energy_balance_sheet.xlsx');

if spider_quick==2
%% 単一データに対するプロット準備
%指標を決定
data_num = input('ボックス・バーチャートで分析したいエネルギー指標の列番号を入力(ex:A→1)：');
fig_title = input('ボックス・バーチャートで出力したいプロットのタイトルは？：');
fig_unit = input('ボックス・バーチャートの出力プロットの単位は？：');
data_num2 = input('スキャッタリングプロットで分析したいもう１つの指標の列番号を入力(ex:A→1)：');
fig_title2 = input('スキャッタリングプロットで出力したいプロットの指標名は？：');
fig_unit2 = input('スキャッタリングプロットで分析したいもう１つの指標の単位は？：');
fig_title_top = input('スキャッタリングプロットで出力したいプロットのタイトルは？：');
axis_choice = input('スキャッタリングプロットの横軸はどっち？(箱ひげ図の指標；1,もう１つの指標；2)：');

%国名/地域/１つめの指標/2つめの指標を格納
a=1; %配列入力のカウンター
for i =12:156 %国リストの区間分ループ
    if ismissing(DATA{i,data_num})==1 || ismissing(DATA{i,data_num2})==1 %データ欠損値ならその国はスキップ
        continue
    end
choosed_data(a,1:4)=[DATA(i,1) DATA(i,2) DATA(i,data_num) DATA(i,data_num2)];
a=a+1;
end
%棒グラフ用
box=1; %配列入力のカウンター 
for i =12:156 %国リストの区間分ループ
     if ismissing(DATA{i,data_num})==1
     continue
     end
choosed_data_box(box,1:2)=[DATA(i,1) DATA(i,data_num)];
box=box+1;
end
%% 通常地域
%%地域別の配列カウンターを作成
europe=1;
africa=1;
middle_east=1;
southeast_asia=1;
south_america=1;
north_america=1;
central_asia_and_the_caucasus=1;
central_america_and_the_caribbean=1;
east_asia=1;
south_asia=1;

%該当地域に分析対象国が存在しない場合のエラー処理のために0を格納しておく
North_America_data={'0' 0 0 0};
Europe_data={'0' 0 0 0};
Africa_data={'0' 0 0 0};
Middle_east_data={'0' 0 0 0};
Southeast_Asia_data={'0' 0 0 0};
South_America_data={'0' 0 0 0};
Central_Asia_and_the_Caucasus_data={'0' 0 0 0};
Central_America_and_the_Caribbean_data={'0' 0 0 0};
East_Asia_data={'0' 0 0 0};
South_Asia_data={'0' 0 0 0};

for i=1:numel(choosed_data(:,1))
if strcmp(choosed_data{i,2},'Europe')==1
    Europe_data(europe,:)=choosed_data(i,:);
    europe=europe+1;
end
if strcmp(choosed_data{i,2},'Africa')==1
    Africa_data(africa,:)=choosed_data(i,:);
    africa=africa+1;
end
if strcmp(choosed_data{i,2},'Middle East')==1
    Middle_east_data(middle_east,:)=choosed_data(i,:);
    middle_east=middle_east+1;
end
if strcmp(choosed_data{i,2},'Southeast Asia')==1
    Southeast_Asia_data(southeast_asia,:)=choosed_data(i,:);
    southeast_asia=southeast_asia+1;
end
if strcmp(choosed_data{i,2},'South America')==1
    South_America_data(south_america,:)=choosed_data(i,:);
    south_america=south_america+1;
end
if strcmp(choosed_data{i,2},'North America')==1
    North_America_data(north_america,:)=choosed_data(i,:);
    north_america=north_america+1;
end
if strcmp(choosed_data{i,2},'Central Asia and the Caucasus')==1
    Central_Asia_and_the_Caucasus_data(central_asia_and_the_caucasus,:)=choosed_data(i,:);
    central_asia_and_the_caucasus=central_asia_and_the_caucasus+1;
end
if strcmp(choosed_data{i,2},'Central America and the Caribbean')==1
    Central_America_and_the_Caribbean_data(central_america_and_the_caribbean,:)=choosed_data(i,:);
    central_america_and_the_caribbean=central_america_and_the_caribbean+1;
end
if strcmp(choosed_data{i,2},'East Asia')==1
    East_Asia_data(east_asia,:)=choosed_data(i,:);
    east_asia=east_asia+1;
end
if strcmp(choosed_data{i,2},'South Asia')==1
    South_Asia_data(south_asia,:)=choosed_data(i,:);
    south_asia=south_asia+1;
end
end

%% High income
%%地域別の配列カウンターを作成
europe_h=1;
africa_h=1;
middle_east_h=1;
southeast_asia_h=1;
south_america_h=1;
north_america_h=1;
central_asia_and_the_caucasus_h=1;
central_america_and_the_caribbean_h=1;
east_asia_h=1;
south_asia_h=1;

%該当地域に分析対象国が存在しない場合のエラー処理のために0を格納しておく
North_America_data_h={'0' 0 0 0};
Europe_data_h={'0' 0 0 0};
Africa_data_h={'0' 0 0 0};
Middle_east_data_h={'0' 0 0 0};
Southeast_Asia_data_h={'0' 0 0 0};
South_America_data_h={'0' 0 0 0};
Central_Asia_and_the_Caucasus_data_h={'0' 0 0 0};
Central_America_and_the_Caribbean_data_h={'0' 0 0 0};
East_Asia_data_h={'0' 0 0 0};
South_Asia_data_h={'0' 0 0 0};

for i=1:numel(choosed_data(:,1))
if strcmp(choosed_data{i,2},'Europe_highincome')==1
    Europe_data_h(europe_h,:)=choosed_data(i,:);
    europe_h=europe_h+1;
end
if strcmp(choosed_data{i,2},'Africa_highincome')==1
    Africa_data_h(africa_h,:)=choosed_data(i,:);
    africa_h=africa_h+1;
end
if strcmp(choosed_data{i,2},'Middle East_highincome')==1
    Middle_east_data_h(middle_east_h,:)=choosed_data(i,:);
    middle_east_h=middle_east_h+1;
end
if strcmp(choosed_data{i,2},'Southeast Asia_highincome')==1
    Southeast_Asia_data_h(southeast_asia_h,:)=choosed_data(i,:);
    southeast_asia_h=southeast_asia_h+1;
end
if strcmp(choosed_data{i,2},'South America_highincome')==1
    South_America_data_h(south_america_h,:)=choosed_data(i,:);
    south_america_h=south_america_h+1;
end
if strcmp(choosed_data{i,2},'North America_highincome')==1
    North_America_data_h(north_america_h,:)=choosed_data(i,:);
    north_america_h=north_america_h+1;
end
if strcmp(choosed_data{i,2},'Central Asia and the Caucasus_highincome')==1
    Central_Asia_and_the_Caucasus_data_h(central_asia_and_the_caucasus_h,:)=choosed_data(i,:);
    central_asia_and_the_caucasus_h=central_asia_and_the_caucasus_h+1;
end
if strcmp(choosed_data{i,2},'Central America and the Caribbean_highincome')==1
    Central_America_and_the_Caribbean_data_h(central_america_and_the_caribbean_h,:)=choosed_data(i,:);
    central_america_and_the_caribbean_h=central_america_and_the_caribbean_h+1;
end
if strcmp(choosed_data{i,2},'East Asia_highincome')==1
    East_Asia_data_h(east_asia_h,:)=choosed_data(i,:);
    east_asia_h=east_asia_h+1;
end
if strcmp(choosed_data{i,2},'South Asia_highincome')==1
    South_Asia_data_h(south_asia_h,:)=choosed_data(i,:);
    south_asia_h=south_asia_h+1;
end
end

%% 地域別の配列から箱ひげ図を作成

%箱ひげ図作成のためにプロットデータを１つにまとめる
region_box=[Europe_data_h(:,3);Africa_data_h(:,3);Middle_east_data_h(:,3);Southeast_Asia_data_h(:,3);Central_Asia_and_the_Caucasus_data_h(:,3);East_Asia_data_h(:,3);South_Asia_data_h(:,3);North_America_data_h(:,3);South_America_data_h(:,3);Central_America_and_the_Caribbean_data_h(:,3);...
    Europe_data(:,3);Africa_data(:,3);Middle_east_data(:,3);Southeast_Asia_data(:,3);Central_Asia_and_the_Caucasus_data(:,3);East_Asia_data(:,3);South_Asia_data(:,3);North_America_data(:,3);South_America_data(:,3);Central_America_and_the_Caribbean_data(:,3)];
%箱ひげ図用にカテゴライズ
g1 = repmat({1},numel(Europe_data_h(:,1)),1);
g2 = repmat({2},numel(Africa_data_h(:,1)),1);
g3 = repmat({3},numel(Middle_east_data_h(:,1)),1);
g4 = repmat({4},numel(Southeast_Asia_data_h(:,1)),1);
g5 = repmat({5},numel(Central_Asia_and_the_Caucasus_data_h(:,1)),1);
g6 = repmat({6},numel(East_Asia_data_h(:,1)),1);
g7 = repmat({7},numel(South_Asia_data_h(:,1)),1);
g8 = repmat({8},numel(North_America_data_h(:,1)),1);
g9 = repmat({9},numel(South_America_data_h(:,1)),1);
g10= repmat({10},numel(Central_America_and_the_Caribbean_data_h(:,1)),1);

g11 = repmat({11},numel(Europe_data(:,1)),1);
g12 = repmat({12},numel(Africa_data(:,1)),1);
g13 = repmat({13},numel(Middle_east_data(:,1)),1);
g14 = repmat({14},numel(Southeast_Asia_data(:,1)),1);
g15 = repmat({15},numel(Central_Asia_and_the_Caucasus_data(:,1)),1);
g16 = repmat({16},numel(East_Asia_data(:,1)),1);
g17 = repmat({17},numel(South_Asia_data(:,1)),1);
g18 = repmat({18},numel(North_America_data(:,1)),1);
g19 = repmat({19},numel(South_America_data(:,1)),1);
g20= repmat({20},numel(Central_America_and_the_Caribbean_data(:,1)),1);

g = [g1; g2; g3; g4; g5; g6; g7; g8; g9; g10; g11; g12; g13; g14; g15; g16; g17; g18; g19; g20];
%プロット用にcellを数値配列に変換
region_box2=cell2mat(region_box);

%箱ひげ図プロット
boxplot(region_box2,g,'Labels',{['Europe' char(10) '(high income)'],['Africa' char(10) '(high income)'],['Middle east' char(10) '(high income)'],['Southeast Asia' char(10) '(high income)'],['Central Asia' char(10) 'and the Caucasus' char(10) '(high income)'],['East Asia' char(10) '(high income)'],['South Asia' char(10) '(high income)'],['North America ' char(10) '(high income)'],['South America' char(10) '(high income)'],['Central America' char(10) 'and the Caribbean' char(10) '(high income)'],...
    'Europe','Africa','Middle east','Southeast Asia',['Central Asia' char(10) 'and the Caucasus'],'East Asia','South Asia','North America','South America',['Central America' char(10) 'and the Caribbean']},'LabelOrientation','inline','FactorSeparator',1,'ColorGroup',[1,1,1,1,1,1,1,1,1,1,2,2,2,2,2,2,2,2,2,2])
hold on
yline(0);
hold on
xline(10.5);
grid on
title([fig_title fig_unit])
ylabel([fig_title fig_unit])
xlabel('Region')
dim1 = [.2 .5 .3 .4];
str1 = 'High income';
annotation('textbox',dim1,'String',str1,'FitBoxToText','on');
dim2 = [.8 .5 .3 .4];
str2 = 'Low~Middle income';
annotation('textbox',dim2,'String',str2,'FitBoxToText','on');
%保存時に最大化画像を出力
h = gcf;
h.WindowState = 'maximized'
%画像を自動保存
saveas(gcf,['C:\Users\Akama\Desktop\works\internship\JICA\IEAデータベース\ビジュアライズ\' fig_title], 'png');
close all
%% 分析対象指標について、世界トップ10, ワースト10, 各地域ごとのランキングの棒グラフを作成。
%% 世界トップ・ワースト10ヶ国の棒グラフ
%全ての国を分析対象指標の大きさ順にソート
world_sort_data=sortrows(choosed_data_box,2);
%トップ・ワースト10ヶ国の国名の配列を作成
country_top=world_sort_data([1:10 numel(world_sort_data(:,1))-9:numel(world_sort_data(:,1))],1);
%トップ・ワースト10ヶ国の指標の数値配列を作成
value_top=world_sort_data([1:10 numel(world_sort_data(:,1))-9:numel(world_sort_data(:,1))],2);
%数値配列をセルから通常配列に変換
value_top2=cell2mat(value_top);
country_list = categorical(country_top);
country_list = reordercats(country_list,country_top);
%プロット
b=bar(country_list,value_top2)
hold on
%世界平均を追加
yl=yline(DATA{10,data_num},'-','World Average','LineWidth',3)
yl.LabelHorizontalAlignment = 'left';
grid on
b.FaceColor = 'flat';
top_c=[.5 0 .5];
top_cc=repmat(top_c,10,1);
b.CData(1:10,:) = top_cc;
worst_c=[255 255 0];
worst_cc=repmat(worst_c,10,1);
b.CData(11:20,:) = worst_cc;
title([fig_title fig_unit])
ylabel([fig_title fig_unit])
xlabel('Country')
dim1 = [.2 .5 .3 .4];
str1 = 'Worst 10 Country';
annotation('textbox',dim1,'String',str1,'FitBoxToText','on');
dim2 = [.8 .5 .3 .4];
str2 = 'Top 10 Country';
annotation('textbox',dim2,'String',str2,'FitBoxToText','on');
%保存時に最大化画像を出力
h = gcf;
h.WindowState = 'maximized'
%画像を自動保存
saveas(gcf,['C:\Users\Akama\Desktop\works\internship\JICA\IEAデータベース\ビジュアライズ\' fig_title '_bar_world'], 'png');
close all


%% 上記の分析指標に加え、もう１つの分析指標を用いて散布図を作成する (地域別クラスタリングver)
%各地域のcell配列を通常配列にして、数値だけ抽出
North_America_data_v=cell2mat(North_America_data(:,[3 4]));
Europe_data_v=cell2mat(Europe_data(:,[3 4]));
Africa_data_v=cell2mat(Africa_data(:,[3 4]));
Middle_east_data_v=cell2mat(Middle_east_data(:,[3 4]));
Southeast_Asia_data_v=cell2mat(Southeast_Asia_data(:,[3 4]));
South_America_data_v=cell2mat(South_America_data(:,[3 4]));
Central_Asia_and_the_Caucasus_data_v=cell2mat(Central_Asia_and_the_Caucasus_data(:,[3 4]));
Central_America_and_the_Caribbean_data_v=cell2mat(Central_America_and_the_Caribbean_data(:,[3 4]));
East_Asia_data_v=cell2mat(East_Asia_data(:,[3 4]));
South_Asia_data_v=cell2mat(South_Asia_data(:,[3 4]));
North_America_data_v_h=cell2mat(North_America_data_h(:,[3 4]));
Europe_data_v_h=cell2mat(Europe_data_h(:,[3 4]));
Africa_data_v_h=cell2mat(Africa_data_h(:,[3 4]));
Middle_east_data_v_h=cell2mat(Middle_east_data_h(:,[3 4]));
Southeast_Asia_data_v_h=cell2mat(Southeast_Asia_data_h(:,[3 4]));
South_America_data_v_h=cell2mat(South_America_data_h(:,[3 4]));
Central_Asia_and_the_Caucasus_data_v_h=cell2mat(Central_Asia_and_the_Caucasus_data_h(:,[3 4]));
Central_America_and_the_Caribbean_data_v_h=cell2mat(Central_America_and_the_Caribbean_data_h(:,[3 4]));
East_Asia_data_v_h=cell2mat(East_Asia_data_h(:,[3 4]));
South_Asia_data_v_h=cell2mat(South_Asia_data_h(:,[3 4]));
%プロットする軸の決め方に応じて各地域のプロット配列を入力
if axis_choice ==1
plotData1=[North_America_data_v(:,1) North_America_data_v(:,2)];
plotData2=[Europe_data_v(:,1) Europe_data_v(:,2)];
plotData3=[Africa_data_v(:,1) Africa_data_v(:,2)];
plotData4=[Middle_east_data_v(:,1) Middle_east_data_v(:,2)];
plotData5=[Southeast_Asia_data_v(:,1) Southeast_Asia_data_v(:,2)];
plotData6=[South_America_data_v(:,1) South_America_data_v(:,2)];
plotData7=[Central_Asia_and_the_Caucasus_data_v(:,1) Central_Asia_and_the_Caucasus_data_v(:,2)];
plotData8=[Central_America_and_the_Caribbean_data_v(:,1) Central_America_and_the_Caribbean_data_v(:,2)];
plotData9=[East_Asia_data_v(:,1) East_Asia_data_v(:,2)];
plotData10=[South_Asia_data_v(:,1) South_Asia_data_v(:,2)];
plotData1_h=[North_America_data_v_h(:,1) North_America_data_v_h(:,2)];
plotData2_h=[Europe_data_v_h(:,1) Europe_data_v_h(:,2)];
plotData3_h=[Africa_data_v_h(:,1) Africa_data_v_h(:,2)];
plotData4_h=[Middle_east_data_v_h(:,1) Middle_east_data_v_h(:,2)];
plotData5_h=[Southeast_Asia_data_v_h(:,1) Southeast_Asia_data_v_h(:,2)];
plotData6_h=[South_America_data_v_h(:,1) South_America_data_v_h(:,2)];
plotData7_h=[Central_Asia_and_the_Caucasus_data_v_h(:,1) Central_Asia_and_the_Caucasus_data_v_h(:,2)];
plotData8_h=[Central_America_and_the_Caribbean_data_v_h(:,1) Central_America_and_the_Caribbean_data_v_h(:,2)];
plotData9_h=[East_Asia_data_v_h(:,1) East_Asia_data_v_h(:,2)];
plotData10_h=[South_Asia_data_v_h(:,1) South_Asia_data_v_h(:,2)];
elseif axis_choice==2
plotData1=[North_America_data_v(:,2) North_America_data_v(:,1)];
plotData2=[Europe_data_v(:,2) Europe_data_v(:,1)];
plotData3=[Africa_data_v(:,2) Africa_data_v(:,1)];
plotData4=[Middle_east_data_v(:,2) Middle_east_data_v(:,1)];
plotData5=[Southeast_Asia_data_v(:,2) Southeast_Asia_data_v(:,1)];
plotData6=[South_America_data_v(:,2) South_America_data_v(:,1)];
plotData7=[Central_Asia_and_the_Caucasus_data_v(:,2) Central_Asia_and_the_Caucasus_data_v(:,1)];
plotData8=[Central_America_and_the_Caribbean_data_v(:,2) Central_America_and_the_Caribbean_data_v(:,1)];
plotData9=[East_Asia_data_v(:,2) East_Asia_data_v(:,1)];
plotData10=[South_Asia_data_v(:,2) South_Asia_data_v(:,1)];
plotData1_h=[North_America_data_v_h(:,2) North_America_data_v_h(:,1)];
plotData2_h=[Europe_data_v_h(:,2) Europe_data_v_h(:,1)];
plotData3_h=[Africa_data_v_h(:,2) Africa_data_v_h(:,1)];
plotData4_h=[Middle_east_data_v_h(:,2) Middle_east_data_v_h(:,1)];
plotData5_h=[Southeast_Asia_data_v_h(:,2) Southeast_Asia_data_v_h(:,1)];
plotData6_h=[South_America_data_v_h(:,2) South_America_data_v_h(:,1)];
plotData7_h=[Central_Asia_and_the_Caucasus_data_v_h(:,2) Central_Asia_and_the_Caucasus_data_v_h(:,1)];
plotData8_h=[Central_America_and_the_Caribbean_data_v_h(:,2) Central_America_and_the_Caribbean_data_v_h(:,1)];
plotData9_h=[East_Asia_data_v_h(:,2) East_Asia_data_v_h(:,1)];
plotData10_h=[South_Asia_data_v_h(:,2) South_Asia_data_v_h(:,1)];
end
%各地域の国名リストを作成
%なし
%各地域をプロット、hold on、色分け、high incomeはドット塗りつぶし
p1=scatter(plotData1(:,1),plotData1(:,2),81,'filled')
text(plotData1(:,1),plotData1(:,2),North_America_data(:,1))
hold on
p2=scatter(plotData2(:,1),plotData2(:,2),81,'filled')
text(plotData2(:,1),plotData2(:,2),Europe_data(:,1))
hold on
p3=scatter(plotData3(:,1),plotData3(:,2),81,'filled')
text(plotData3(:,1),plotData3(:,2),Africa_data(:,1))
hold on
p4=scatter(plotData4(:,1),plotData4(:,2),81,'filled')
text(plotData4(:,1),plotData4(:,2),Middle_east_data(:,1))
hold on
p5=scatter(plotData5(:,1),plotData5(:,2),81,'filled')
text(plotData5(:,1),plotData5(:,2),Southeast_Asia_data(:,1))
hold on
p6=scatter(plotData6(:,1),plotData6(:,2),81,'filled')
text(plotData6(:,1),plotData6(:,2),South_America_data(:,1))
hold on
p7=scatter(plotData7(:,1),plotData7(:,2),81,'filled')
text(plotData7(:,1),plotData7(:,2),Central_Asia_and_the_Caucasus_data(:,1))
hold on
p8=scatter(plotData8(:,1),plotData8(:,2),81,'filled','k')
text(plotData8(:,1),plotData8(:,2),Central_America_and_the_Caribbean_data(:,1))
hold on
p9=scatter(plotData9(:,1),plotData9(:,2),81,[1 0.078 0.57],'filled')
text(plotData9(:,1),plotData9(:,2),East_Asia_data(:,1))
hold on
p10=scatter(plotData10(:,1),plotData10(:,2),81,[0.67 1 0.18],'filled')
text(plotData10(:,1),plotData10(:,2),South_Asia_data(:,1))
hold on
p11=scatter(plotData1_h(:,1),plotData1_h(:,2),81)
text(plotData1_h(:,1),plotData1_h(:,2),North_America_data_h(:,1))
hold on
p12=scatter(plotData2_h(:,1),plotData2_h(:,2),81)
text(plotData2_h(:,1),plotData2_h(:,2),Europe_data_h(:,1))
hold on
p13=scatter(plotData3_h(:,1),plotData3_h(:,2),81)
text(plotData3_h(:,1),plotData3_h(:,2),Africa_data_h(:,1))
hold on
p14=scatter(plotData4_h(:,1),plotData4_h(:,2),81)
text(plotData4_h(:,1),plotData4_h(:,2),Middle_east_data_h(:,1))
hold on
p15=scatter(plotData5_h(:,1),plotData5_h(:,2),81)
text(plotData5_h(:,1),plotData5_h(:,2),Southeast_Asia_data_h(:,1))
hold on
p16=scatter(plotData6_h(:,1),plotData6_h(:,2),81)
text(plotData6_h(:,1),plotData6_h(:,2),South_America_data_h(:,1))
hold on
p17=scatter(plotData7_h(:,1),plotData7_h(:,2),81)
text(plotData7_h(:,1),plotData7_h(:,2),Central_Asia_and_the_Caucasus_data_h(:,1))
hold on
p18=scatter(plotData8_h(:,1),plotData8_h(:,2),81,'k')
text(plotData8_h(:,1),plotData8_h(:,2),Central_America_and_the_Caribbean_data_h(:,1))
hold on
p19=scatter(plotData9_h(:,1),plotData9_h(:,2),81,[1 0.078 0.57])
text(plotData9_h(:,1),plotData9_h(:,2),East_Asia_data_h(:,1))
hold on
p20=scatter(plotData10_h(:,1),plotData10_h(:,2),81,[0.67 1 0.18])
text(plotData10_h(:,1),plotData10_h(:,2),South_Asia_data_h(:,1))
hold on
%世界平均を追加
if axis_choice==1
yl=yline(DATA{10,data_num2},'-','World Average','LineWidth',3)
yl.LabelHorizontalAlignment = 'right';
hold on
xl=xline(DATA{10,data_num},'-','World Average','LineWidth',3)
xl.LabelHorizontalAlignment = 'left';
elseif axis_choice==2
yl=yline(DATA{10,data_num},'-','World Average','LineWidth',3)
yl.LabelHorizontalAlignment = 'right';
hold on
xl=xline(DATA{10,data_num2},'-','World Average','LineWidth',3)
xl.LabelHorizontalAlignment = 'left';
end
grid on;
lgd = legend([p1,p2,p3,p4,p5,p6,p7,p8,p9,p10,p11,p12,p13,p14,p15,p16,p17,p18,p19,p20,xl,yl],...
    { 'North America','Europe','Africa','Middle east','Southeast Asia','South America',['Central Asia' char(10) 'and the Caucasus'],['Central America' char(10) 'and the Caribbean'],'East Asia','South Asia',...
     ['North America ' char(10) '(high income)'],['Europe' char(10) '(high income)'],['Africa' char(10) '(high income)'],['Middle east' char(10) '(high income)'],['Southeast Asia' char(10) '(high income)'],...
    ['South America' char(10) '(high income)'],['Central Asia' char(10) 'and the Caucasus' char(10) '(high income)'],['Central America' char(10) 'and the Caribbean' char(10) '(high income)'],...
    ['East Asia' char(10) '(high income)'],['South Asia' char(10) '(high income)'],...
    'World Average(x)','World Average(y)'},'NumColumns',2);
title(lgd,'Region Classification')
% xlim([0 max(plotData(:,1))])
% ylim([0 max(plotData(:,2))])
title(fig_title_top)
if axis_choice==1
ylabel([fig_title2 fig_unit2])
xlabel([fig_title fig_unit])
elseif axis_choice==2
ylabel([fig_title fig_unit])
xlabel([fig_title2 fig_unit2])
end
%保存時に最大化画像を出力
h = gcf;
h.WindowState = 'maximized'
%画像を自動保存
saveas(gcf,['C:\Users\Akama\Desktop\works\internship\JICA\IEAデータベース\ビジュアライズ\' fig_title '_scatter_world'], 'png');
% set(gca,'FontName','arial','FontSize',12)
% saveas(gca,'usage_02','png')

fig_close = input('散布図のfigureウィンドウを閉じますか？(好きな縮尺で保存し終わったら：１)：');
if fig_close==1
    close all
end
end
disp('////////////// ここからレーダチャートを作っていくよ！////////////////////')
%% レーダーチャート
% 仕様：散布図まで出力されたら、散布図のfigureウィンドウを閉じずに2つ目のウィンドウ上でレーダチャートを出力。
% 　　　地域名をインプットさせ、地域別データ配列に格納する５つの指標を入力。
%     　forループで1か国ごとにfigureを出力、PCに自動保存
%MATLAB2020bにはレーダチャートがデフォルト関数として存在しないのでFile Exchangeから取得
% https://jp.mathworks.com/matlabcentral/fileexchange/59561-spider_plot

%地域を選択
region_choice = input('どの地域を分析する？：');
%分析したいデータ指標を５つ選択
data_choice1 = input('分析する指標は？（1つ目/Excel列番号）：');
data_choice2 = input('分析する指標は？（2つ目/Excel列番号）：');
data_choice3 = input('分析する指標は？（3つ目/Excel列番号）：');
data_choice4 = input('分析する指標は？（4つ目/Excel列番号）：');
data_choice5 = input('分析する指標は？（5つ目/Excel列番号）：');
%分析したいデータ指標の名前を取得
data_label1=DATA{6,data_choice1};
data_label2=DATA{6,data_choice2};
data_label3=DATA{6,data_choice3};
data_label4=DATA{6,data_choice4};
data_label5=DATA{6,data_choice5};

%該当地域に属する国名配列と各国５つずつのデータ配列を作成。
a=1; %配列入力のカウンター
for i =12:156 %国リストの数だけループ
%データ欠損値があればその国はスキップ,該当地域なら配列作成
if strcmp(DATA{i,2},region_choice)==1 && (ismissing(DATA{i,data_choice1})==0 && ismissing(DATA{i,data_choice2})==0 && ismissing(DATA{i,data_choice3})==0 && ismissing(DATA{i,data_choice4})==0 && ismissing(DATA{i,data_choice5})==0 )
    country_list_s{a,1}=DATA{i,1};
    choosed_data_s(a,1:5)=[DATA{i,data_choice1} DATA{i,data_choice2} DATA{i,data_choice3} DATA{i,data_choice4} DATA{i,data_choice5}];
    a=a+1;
end
end
%データ配列の最後の2行に世界平均と地域平均の行を追加
%地域平均
choosed_data_s(numel(choosed_data_s(:,1))+1,1:5)=[mean(choosed_data_s(:,1)) mean(choosed_data_s(:,2)) mean(choosed_data_s(:,3)) mean(choosed_data_s(:,4)) mean(choosed_data_s(:,5))];
%世界平均
choosed_data_s(numel(choosed_data_s(:,1))+1,1:5)=[DATA{10,data_choice1} DATA{10,data_choice2} DATA{10,data_choice3} DATA{10,data_choice4} DATA{10,data_choice5}];

%各国ずつレーダチャートを作成し保存していく。
for i=1:numel(country_list_s(:,1))
    P=[choosed_data_s(i,:); choosed_data_s(numel(choosed_data_s(:,1))-1,:); choosed_data_s(numel(choosed_data_s(:,1)),:)]
    s = spider_plot_class(P);
    %レーダチャートのレイアウト調整
    %レーダ線の凡例
    s.LegendLabels = {country_list_s{i,1}, 'redional average', 'world average'};
    s.LegendHandle.Location = 'northeastoutside';
    %レーダーのラベル
    s.AxesLabels = {data_label1, data_label2, data_label3, data_label4, data_label5};
    s.AxesInterval = 2;
    %レーダーの塗りつぶし
    s.FillOption = {'on', 'on', 'off'};
    s.FillTransparency = [0.2, 0.1, 0.1];
    %レーダ値のリミット設定
    %s.AxesLimits = [1, 2, 1, 1, 1; 10, 8, 9, 5, 10]; % [min axes limits; max axes limits]
    %レーダ値の有効数字
    s.AxesPrecision = [3, 3, 3, 3, 3];
    %保存時に最大化画像を出力
    h = gcf;
    h.WindowState = 'maximized'
    %画像を保存
    saveas(gcf,['C:\Users\Akama\Desktop\works\internship\JICA\IEAデータベース\ビジュアライズ\' country_list_s{i,1} '_spider_plot'], 'png');
    %figureウィンドウを閉じる
    close all
end

%% ここから、選択した２か国について１つのレーダーチャートを作成
%比較する２か国を入力
country_choice1 = input('どの国を比較する？（1か国目）：');
country_choice2 = input('どの国を比較する？（2か国目）：');
%入力した2か国が配列の何行目にあるか検出
for i=1:numel(country_list_s(:,1))
    if isequal(country_list_s{i,1},country_choice1)==1
        c1_num=i;
    end
    if isequal(country_list_s{i,1},country_choice2)==1
        c2_num=i;
    end
end
%検出した配列の行数の値からレーダーチャートを作成
P2=[choosed_data_s(c1_num,:); choosed_data_s(c2_num,:); choosed_data_s(numel(choosed_data_s(:,1))-1,:); choosed_data_s(numel(choosed_data_s(:,1)),:)]
    s = spider_plot_class(P2);
    %レーダチャートのレイアウト調整
    %レーダ線の凡例
    s.LegendLabels = {country_list_s{c1_num,1}, country_list_s{c2_num,1}, 'redional average', 'world average'};
    s.LegendHandle.Location = 'northeastoutside';
    %レーダーのラベル
    s.AxesLabels = {data_label1, data_label2, data_label3, data_label4, data_label5};
    s.AxesInterval = 2;
    %レーダーの塗りつぶし
    s.FillOption = {'on', 'on', 'off', 'off'};
    s.FillTransparency = [0.2, 0.2, 0.1, 0.1];
    %レーダ値のリミット設定
    s.AxesLimits = [0, 0, 0, 0, 0; 100, 100, 100, 100, 100]; % [min axes limits; max axes limits]
    %レーダ値の有効数字
    s.AxesPrecision = [3, 3, 3, 3, 3];
    %保存時に最大化画像を出力
    h = gcf;
    h.WindowState = 'maximized'
    %画像を保存
    saveas(gcf,['C:\Users\Akama\Desktop\works\internship\JICA\IEAデータベース\ビジュアライズ\' country_list_s{c1_num,1} '_' country_list_s{c1_num,1} '_spider_plot_compare'], 'png');
    %figureウィンドウを閉じる
    close all
    
    %% Excel統合データのアップロード時のコマンド
% T=readmatrix();%ここにIEAの元データを読み込み
% filename = 'energy_balance_sheet.xlsx';　%ここで、更新して上書き保存するExcel統合ファイルの名前を入力
% writematrix(T,filename,'Range','BG12:BH12') %更新するセルの範囲を指定して書き込み開始

%WB Data BankやBNEFなどの国リストがIEAと異なるデータの場合
% DATA=readcell('energy_balance_sheet.xlsx'); %更新したい統合エクセルファイルを読み込み
% DATA_new=readcell('data.xlsx'); %更新するのに用いる新しいデータを読み込み
%c_line=n; %nに国リストの最初の行を入力（例：最初のAlbaniaが12行目なら12）
%c_line2=m; %nに国リストの最後の行を入力（例：最後のzimbabweが145行目なら145）
%for i=c_line:c_line2 %元の国リストの数だけ繰り返す
%a=0; %配列のカウンター
% for j=c_line:c_line2 %新しいデータの方の国リストの数だけ繰り返す
% if DATA{i,1}==DATA2{j,1}
%     DATA_save(i,:)=DATA2{j,2:length(DATA2)};
%     a=a+1
% end
% end
%DATA_save2=cell2mat(DATA_save);
% filename = 'energy_balance_sheet.xlsx';　%ここで、更新して上書き保存するExcel統合ファイルの名前を入力
% writematrix(T,filename,'Range','BG12:BH12') %更新するセルの範囲を指定して書き込み開始

%% １つの国を選択して、すべての指標について、世界平均に対する割合を棒グラフで出力
%分析する国を選択
country_name=input('どの国を分析する？：');
%その国がエクセルの何行目にあるかを検索
for i=1:numel(DATA(:,1))
    if isequal(DATA{i,1},country_name)==1
        c_num=i;
    end
end
%分析したい指標のエクセル列番号を全て入力
data_list=input('分析したい指標の列番号を配列[]で入力してね：');
%分析したい指標の名前を取得　
a=1;
for i=data_list
    data_name{a,1}=DATA{6,i};
    a=a+1;
end
%分析したい国のデータと世界平均に対する割合を配列に格納 %平均に対する偏差では指標ごとの差が大きくなってしますためボツ
a=1;
for i=data_list
    if ismissing(DATA{c_num,i})==1
        DATA{c_num,i}=0;
    end
    data_set(a,1)=DATA{c_num,i}/DATA{10,i};
    a=a+1;
end
%セルを配列に変換
data_name=categorical(data_name);
%プロット
b=bar(data_name,data_set)
hold on
grid on
b.FaceColor = 'flat';
title([country_name 'のエネルギー指標の世界平均値に対する割合'])
ylabel('世界平均値に対する割合')
xlabel('DATA')
dim1 = [.2 .5 .3 .4];
str1 = country_name;
annotation('textbox',dim1,'String',str1,'FitBoxToText','on');
%世界平均を追加
yl=yline(1,'-','世界平均値','LineWidth',3)
yl.LabelHorizontalAlignment = 'right';
hold on
%保存時に最大化画像を出力
h = gcf;
h.WindowState = 'maximized'
%画像を自動保存
saveas(gcf,['C:\Users\Akama\Desktop\works\internship\JICA\IEAデータベース\ビジュアライズ\' country_name '_差分'], 'png');
close all
%% 終了処理
disp('分析は正常に完了しました (*‘∀‘)ﾉｼ/// by インターン赤間')