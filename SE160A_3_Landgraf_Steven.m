close all
clear all
clc
format compact
inFile = 'SE160A_3_StressStrain_Input.xlsx';
outFile = 'SE160A_3_StressStrain_Output.xlsx';

xlswrite(outFile, {'Steven Landgraf'}, 1, 'C4');
xlswrite(outFile, {'A10850070'}  , 1, 'C5');
%Input Read and Echo
[~,titles] = xlsread(inFile,1, 'C4');
xlswrite(outFile,titles,1,'C7');
%Part 1 inputs
xx2d = xlsread(inFile, 1, 'C9');
xlswrite(outFile,xx2d,1,'C15');

yy2d = xlsread(inFile, 1, 'C10');
xlswrite(outFile,yy2d,1,'C16');

t2d = xlsread(inFile, 1, 'C11');
xlswrite(outFile,t2d,1,'C17');

alpha1 = xlsread(inFile, 1, 'C12');
xlswrite(outFile,alpha1,1,'C18');
%Part 2 inputs
xx3d = xlsread(inFile, 1, 'C17');
xlswrite(outFile,xx3d,1,'C49');

yy3d = xlsread(inFile, 1, 'C18');
xlswrite(outFile,yy3d,1,'C50');

zz3d = xlsread(inFile, 1, 'C19');
xlswrite(outFile,zz3d,1,'C51');

yz3d = xlsread(inFile, 1, 'C20');
xlswrite(outFile,yz3d,1,'C52');

txz3d = xlsread(inFile, 1, 'C21');
xlswrite(outFile,txz3d,1,'C53');

txy3d = xlsread(inFile, 1, 'C22');
xlswrite(outFile,txy3d,1,'C54');
%Part 3 inputs
exx2d = xlsread(inFile, 1, 'C27');
xlswrite(outFile,exx2d,1,'C73');

eyy2d = xlsread(inFile, 1, 'C28');
xlswrite(outFile,eyy2d,1,'C74');

gxy2d = xlsread(inFile, 1, 'C29');
xlswrite(outFile,gxy2d,1,'C75');

alpha2 = xlsread(inFile, 1, 'C30');
xlswrite(outFile,alpha2,1,'C76');
%Part 4
exx3d = xlsread(inFile,1, 'C35');
xlswrite(outFile,exx3d,1,'C107');

eyy3d = xlsread(inFile,1, 'C36');
xlswrite(outFile,eyy3d,1,'C108');

ezz3d = xlsread(inFile,1, 'C37');
xlswrite(outFile,ezz3d,1,'C109');

gyz3d = xlsread(inFile,1, 'C38');
xlswrite(outFile,gyz3d,1,'C110');

gxz3d = xlsread(inFile,1, 'C39');
xlswrite(outFile,gxz3d,1,'C111');

gxy3d = xlsread(inFile,1, 'C40');
xlswrite(outFile,gxy3d,1,'C112');
%Part 5
sga = xlsread(inFile,1, 'C45');
xlswrite(outFile,sga,1,'C131');

sgao = xlsread(inFile,1, 'D45');
xlswrite(outFile,sgao,1,'D131');

sgb = xlsread(inFile,1, 'C46');
xlswrite(outFile,sgb,1,'C132');

sgbo = xlsread(inFile,1, 'D46');
xlswrite(outFile,sgbo,1,'D132');

sgc = xlsread(inFile,1, 'C47');
xlswrite(outFile,sgc,1,'C133');

sgco = xlsread(inFile,1, 'D47');
xlswrite(outFile,sgco,1,'D133');

alpha3 = xlsread(inFile,1, 'C48');
xlswrite(outFile,alpha3,1,'C134');

%Part 1 Calculations
xxprime=xx2d*(cosd(alpha1)^2)+yy2d*(sind(alpha1)^2)+2*t2d*(cosd(alpha1)*sind(alpha1));
xlswrite(outFile,xxprime,1,'C23');

yyprime=xx2d*(sind(alpha1)^2)+yy2d*(cosd(alpha1)^2)-2*t2d*(cosd(alpha1)*sind(alpha1));
xlswrite(outFile,yyprime,1,'C24');

txyprime=(yy2d-xx2d)*cosd(alpha1)*sind(alpha1)+t2d*((cosd(alpha1)^2)-(sind(alpha1)^2));
xlswrite(outFile,txyprime,1,'C25');

alphap=0.5*atand((2*t2d)/(xx2d-yy2d));
xlswrite(outFile,alphap,1,'C30');

T=[cosd(alphap) sind(alphap);
    -sind(alphap) cosd(alphap)];
A=[xx2d t2d;
    t2d yy2d];
P=T*A*T.';
xlswrite(outFile,P(1,1),1,'C28');
xlswrite(outFile,P(2,2),1,'C29');
[V,D]=eig(A);
xlswrite(outFile,eig(A).',1,'C33:D33');
xlswrite(outFile,V,1,'C34:D35');

alphas=alphap-45;
xlswrite(outFile,alphas,1,'C41');

xxmax=xx2d*(cosd(alphas)^2)+yy2d*(sind(alphas)^2)+2*t2d*(cosd(alphas)*sind(alphas));
xlswrite(outFile,xxmax,1,'C38');

yymax=xx2d*(sind(alphas)^2)+yy2d*(cosd(alphas)^2)-2*t2d*(cosd(alphas)*sind(alphas));
xlswrite(outFile,yymax,1,'C39');

tmax=(yy2d-xx2d)*cosd(alphas)*sind(alphas)+t2d*((cosd(alphas)^2)-(sind(alphas)^2));
xlswrite(outFile,tmax,1,'C40');

%Part 2 Calculations
A=[xx3d txy3d txz3d;
    txy3d yy3d yz3d;
    txz3d yz3d zz3d];
xlswrite(outFile,eig(A).',1,'C59:E59');
[V,D]=eig(A);
xlswrite(outFile,V,1,'C60:E62')
xlswrite(outFile,(max(eig(A))-min(eig(A)))*0.5,1,'C65');

%Part 3 Calculations
exx2dprime=exx2d*(cosd(alpha2)^2)+eyy2d*(sind(alpha2)^2)+gxy2d*(cosd(alpha2)*sind(alpha2));
xlswrite(outFile,exx2dprime,1,'C81');

eyy2dprime=exx2d*(sind(alpha2)^2)+eyy2d*(cosd(alpha2)^2)-gxy2d*(cosd(alpha2)*sind(alpha2));
xlswrite(outFile,eyy2dprime,1,'C82');

gxyprime=2*(eyy2d-exx2d)*cosd(alpha2)*sind(alpha2)+gxy2d*((cosd(alpha2)^2)-(sind(alpha2)^2));
xlswrite(outFile,gxyprime,1,'C83');

alphap=0.5*atand(gxy2d/(exx2d-eyy2d));
xlswrite(outFile,alphap,1,'C88');

E=[exx2d 0.5*gxy2d;
    0.5*gxy2d eyy2d];
exx2dprime=exx2d*(cosd(alphap)^2)+eyy2d*(sind(alphap)^2)+gxy2d*(cosd(alphap)*sind(alphap));
xlswrite(outFile,exx2dprime,1,'C86');

eyy2dprime=exx2d*(sind(alphap)^2)+eyy2d*(cosd(alphap)^2)-gxy2d*(cosd(alphap)*sind(alphap));
xlswrite(outFile,eyy2dprime,1,'C87');
xlswrite(outFile,eig(E).',1,'C91:D91');

[V,D]=eig(E);
xlswrite(outFile,V,1,'C92:D93');

as=(alphap-45);
xlswrite(outFile,as,1,'C99');

exx2dnew=exx2d*(cosd(as)^2)+eyy2d*(sind(as)^2)+gxy2d*(cosd(as)*sind(as));
xlswrite(outFile,exx2dnew,1,'C96');

eyy2dnew=exx2d*(sind(as)^2)+eyy2d*(cosd(as)^2)-gxy2d*(cosd(as)*sind(as));
xlswrite(outFile,eyy2dnew,1,'C97');

gxynew=2*(eyy2d-exx2d)*cosd(as)*sind(as)+gxy2d*((cosd(as)^2)-(sind(as)^2));
xlswrite(outFile,gxynew,1,'C98');

%Part 4 Calculations
E=[exx3d gxy3d/2 gxz3d/2;
    gxy3d/2 eyy3d gyz3d/2;
    gxz3d/2 gyz3d/2 ezz3d];
[V,D]=eig(E);
xlswrite(outFile,max(D),1,'C117:E117');
xlswrite(outFile,V,1,'C118:E120');
xlswrite(outFile,max(eig(E))-min(eig(E)),1,'C123');

%Part 5 Calculations
T=[cosd(sgao)^2 sind(sgao)^2 cosd(sgao)*sind(sgao);
    cosd(sgbo)^2 sind(sgbo)^2 cosd(sgbo)*sind(sgbo);
    cosd(sgco)^2 sind(sgco)^2 cosd(sgco)*sind(sgco)];
P=[sga; sgb; sgc];
E=T^-1*P;
xlswrite(outFile,E,1,'C139:C141');
SG1=E(1)*(cosd(alpha3)^2)+E(2)*(sind(alpha3)^2)+E(3)*(cosd(alpha3)*sind(alpha3));
xlswrite(outFile,SG1,1,'C144');

SG2=E(1)*(sind(alpha3)^2)+E(2)*(cosd(alpha3)^2)-E(3)*(cosd(alpha3)*sind(alpha3));
xlswrite(outFile,SG2,1,'C145');

SG3=2*(E(2)-E(1))*cosd(alpha3)*sind(alpha3)+E(3)*((cosd(alpha3)^2)-(sind(alpha3)^2));
xlswrite(outFile,SG3,1,'C146');

P=[E(1) 0.5*E(3);
    0.5*E(3) E(2)];

xlswrite(outFile,eig(P).',1,'C154:D154');
[V,D]=eig(P);
xlswrite(outFile,V,1,'C155:D156');
alphap=0.5*atand(E(3)/(E(1)-E(2)));
xlswrite(outFile,alphap,1,'C151');

SG1=0.5*(E(1)+E(2))+0.5*(E(1)-E(2))*cosd(2*alphap)+0.5*E(3)*sind(2*alphap);
xlswrite(outFile,SG1,1,'C149');

SG2=0.5*(E(1)+E(2))-0.5*(E(1)-E(2))*cosd(2*alphap)-0.5*E(3)*sind(2*alphap);
xlswrite(outFile,SG2,1,'C150');

alpha3=(alphap-45);
xlswrite(outFile,alpha3,1,'C162');

SG1=E(1)*(cosd(alpha3)^2)+E(2)*(sind(alpha3)^2)+E(3)*(cosd(alpha3)*sind(alpha3));
xlswrite(outFile,SG1,1,'C159');

SG2=E(1)*(sind(alpha3)^2)+E(2)*(cosd(alpha3)^2)-E(3)*(cosd(alpha3)*sind(alpha3));
xlswrite(outFile,SG2,1,'C160');

SG3=2*(E(2)-E(1))*cosd(alpha3)*sind(alpha3)+E(3)*((cosd(alpha3)^2)-(sind(alpha3)^2));
xlswrite(outFile,SG3,1,'C161');


