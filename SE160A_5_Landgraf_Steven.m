close all
clear all
clc
format compact
format long
inFile = 'SE160A_5_Wing_Bending_Input.xlsx';
outFile = 'SE160A_5_Wing_Bending_Output.xlsx';

xlswrite(outFile, {'Steven Landgraf'}, 1, 'C5');
xlswrite(outFile, {'A10850070'}  , 1, 'C6');
%Input Read and Echo
[~,titles] = xlsread(inFile,1, 'C5');
xlswrite(outFile,titles,1,'C8');

%Stringer Definitions
y=xlsread(inFile,1,'C10:F10');
z=xlsread(inFile,1,'C11:F11');
A=xlsread(inFile,1,'C13:F13');
Iyy=xlsread(inFile,1,'C14:F14');
Izz=xlsread(inFile,1,'C15:F15');
Iyz=xlsread(inFile,1,'C16:F16');
E=xlsread(inFile,1,'C18:F18')*10^6;
yield=xlsread(inFile,1,'C19:F19')*10^3;
ultt=xlsread(inFile,1,'C20:F20')*10^3;
ultc=xlsread(inFile,1,'C21:F21')*10^3;

xlswrite(outFile,y,1,'C15:F15');
xlswrite(outFile,z,1,'C16:F16');
xlswrite(outFile,A,1,'C18:F18');
xlswrite(outFile,Iyy,1,'C19:F19');
xlswrite(outFile,Izz,1,'C20:F20');
xlswrite(outFile,Iyz,1,'C21:F21');
xlswrite(outFile,E,1,'C23:F23');
xlswrite(outFile,yield,1,'C24:F24');
xlswrite(outFile,ultt,1,'C25:F25');
xlswrite(outFile,ultc,1,'C26:F26');

%Wing Skin Definitions
t=xlsread(inFile,1,'C26:F26');
G=xlsread(inFile,1,'C27:F27')*10^6;
sheary=xlsread(inFile,1,'C28:F28')*10^3;
shearult=xlsread(inFile,1,'C29:F29')*10^3;

xlswrite(outFile,t,1,'C31:F31');
xlswrite(outFile,G,1,'C32:F32');
xlswrite(outFile,sheary,1,'C33:F33');
xlswrite(outFile,shearult,1,'C34:F34');

%Wing size weight and structural definitions
L=xlsread(inFile,1,'C34');
c=xlsread(inFile,1,'C35');
w=xlsread(inFile,1,'C36');
nu=xlsread(inFile,1,'C37');
FS=xlsread(inFile,1,'C38:C39');

xlswrite(outFile,L,1,'C39');
xlswrite(outFile,c,1,'C40');
xlswrite(outFile,w,1,'C41');
xlswrite(outFile,nu,1,'C42');
xlswrite(outFile,FS,1,'C43:C44');

%Wing Aerodynamic Definition
lift=xlsread(inFile,1,'C44:C46');
drag=xlsread(inFile,1,'C47:C48');
order=xlsread(inFile,1,'C49');
mx=xlsread(inFile,1,'C50');

xlswrite(outFile,lift,1,'C49:C51');
xlswrite(outFile,drag,1,'C52:C53');
xlswrite(outFile,order,1,'C54');
xlswrite(outFile,mx,1,'C55');

EA=0;
EIyy=0;
EIzz=0;
EIyz=0;

for i=1:4
    EA=E(i)*A(i)+EA;
end

yc=0;
zc=0;

for i=1:4
    yc=((E(i)*y(i)*A(i))/EA)+yc;
    zc=((E(i)*z(i)*A(i))/EA)+zc;
end
for i=1:4
    EIyy=(E(i)*((z(i)-zc)^2)*A(i))+EIyy;
    EIzz=(E(i)*((y(i)-yc)^2)*A(i))+EIzz;
    EIyz=(E(i)*(y(i)-yc)*(z(i)-zc)*A(i))+EIyz;
end

a=y(4)-y(1);
b=z(4)-z(1);
peri124=0.25*pi*((3*(a+b))-(((3*a+b)*(a+3*b))^0.5));
if (y(4)-y(1))<(c/2)
    peri234=((c/2)-y(4)+y(1))+((y(3)-y(1)-(c/2))^2+(z(4)-z(3))^2)^0.5;
    Area=(0.5*pi*a*b)+2*b*((c/2)-(y(4)-y(1)))+(c/2)*b;
else
    peri234=((y(3)-y(4))^2+(z(4)-z(3))^2)^0.5;
    Area=(0.5*pi*a*b)+b*(y(3)-y(4));
end
for i=1:4
    GT(i)=G(i)*t(i);
end
GJsum=(peri124/GT(1))+(peri234/GT(2))+(peri234/GT(3))+(peri124/GT(4));

GJ=(4*Area^2)/GJsum;


PX=0;
PY=drag(1)*L+(drag(2)*L/(order+1));
PZ=L*(lift(1)+(lift(2)/3)+(lift(3)/5))-w*nu*L;
MX=mx*L+((c/4)-yc)*L*(lift(1)+(lift(2)/3)+(lift(3)/5))-((c/2)-yc)*w*nu*L;
MY=-(L^2*(0.5*lift(1)+(lift(2)/4)+(lift(3)/6))-0.5*w*nu*L^2);
MZ=(0.5*drag(1)*L^2)+((drag(2)*L^2)/(order+2));

Mat=[EA 0 0;
     0 EIzz EIyz;
     0 EIyz EIyy];
Mat2=[PX;
    MZ;
    -MY];
for i=1:4
    sig(i)=E(i)*[1 -(y(i)-yc) -(z(i)-zc)]*Mat^-1*Mat2;
    if sig(i)<0
        M(i)=min([ultc(i)/(FS(2)*sig(i)) -yield(i)/(FS(1)*sig(i))])-1;
    else
        M(i)=min([ultt(i)/(FS(2)*sig(i)) yield(i)/(FS(1)*sig(i))])-1;
    end
end

ez=0;


syms x q0
rz=lift(1)+lift(2)*(x/L)^2+lift(3)*(x/L)^4;
weight=w*nu;
ry=drag(1)+drag(2)*(x/L)^order;
VZ=int(rz,x,0,L)-int(rz,x,0,x);
WEIGHT=w*nu*L-int(weight,x,0,x);
VY=PY-int(ry,x,0,x);
mex=mx+((c/4)-yc)*rz-((c/2)-yc)*weight;
Mex=VZ*(yc-a)+(MX-int(mex,x,0,x))-WEIGHT*((yc)-a);

A1=(0.25*pi*a*b);
A2=0.5*(2*b*((c/2)-(y(4)-y(1)))+(c/2)*b);
A3=0.5*b*(y(3)-y(4));
q12=q0;
q23=q12-((VY/EIzz)*E(2)*A(2)*(y(2)-yc))-(((VZ-WEIGHT)/EIyy)*E(2)*A(2)*z(2));
q34=q23-((VY/EIzz)*E(3)*A(3)*(y(3)-yc))-(((VZ-WEIGHT)/EIyy)*E(3)*A(3)*z(3));
q41=q34-((VY/EIzz)*E(4)*A(4)*(y(4)-yc))-(((VZ-WEIGHT)/EIyy)*E(4)*A(4)*z(4));
if (y(4)-y(1))<(c/2)
    q0=0;
    Mint=(2*q12*A1)+(2*q41*A1)+(2*q23*A2)+(2*q34*A2);
    Qvar=4*A1+4*A2;
    Mint=subs(Mint);
else
    q0=0;
    Mint=(2*q12*A1)+(2*q23*A3)+(2*q34*A3)+(2*q41*A1);
    Qvar=4*A1+4*A3;
    Mint=subs(Mint);
end
x=0;
Qsolve=Mex-Mint;
Qsolve=subs(Qsolve);
Q=Qsolve/Qvar;
Q=double(Q);
she1=Q/t(1);
q0=Q;
q12=q0;
q23=q12-((VY/EIzz)*E(2)*A(2)*(y(2)-yc))-(((VZ-WEIGHT)/EIyy)*E(2)*A(2)*z(2));
Q23=double(subs(q23));
q34=q23-((VY/EIzz)*E(3)*A(3)*(y(3)-yc))-(((VZ-WEIGHT)/EIyy)*E(3)*A(3)*z(3));
Q34=double(subs(q34));
q41=q34-((VY/EIzz)*E(4)*A(4)*(y(4)-yc))-(((VZ-WEIGHT)/EIyy)*E(4)*A(4)*z(4));
Q41=double(subs(q41));
she2=Q23/t(1);
she3=Q34/t(1);
she4=Q41/t(1);
MS(1)=min([abs(sheary/(FS(1)*she1)) abs(shearult/(FS(2)*she1))])-1;
MS(2)=min([abs(sheary/(FS(1)*she2)) abs(shearult/(FS(2)*she2))])-1;
MS(3)=min([abs(sheary/(FS(1)*she3)) abs(shearult/(FS(2)*she3))])-1;
MS(4)=min([abs(sheary/(FS(1)*she4)) abs(shearult/(FS(2)*she4))])-1;

syms x
MYx=MY+int(VZ-WEIGHT,x,0,x);
MZx=MZ-int(VY,x,0,x);
tipv=double(int(int(MZx,x,0,x),x,0,L))/EIzz;
tipw=-double(int(int(MYx,x,0,x),x,0,L))/EIyy;

flow1=q12*peri124/(G(2)*t(2));
flow2=q23*peri234/(G(3)*t(3));
flow3=q34*peri234/(G(4)*t(4));
flow4=q41*peri124/(G(1)*t(1));
flowsum=flow1+flow2+flow3+flow4;
twistrate=flowsum/(2*Area);
twist=double(int(twistrate,x,0,L))*180/pi;
twists=int(twistrate,x,0,x)==0;
center=solve(twists);

xlswrite(outFile,yc,1,'C62');
xlswrite(outFile,zc,1,'C63');
xlswrite(outFile,EA,1,'C64');
xlswrite(outFile,EIyy,1,'C65');
xlswrite(outFile,EIzz,1,'C66');
xlswrite(outFile,EIyz,1,'C67');
xlswrite(outFile,GJ,1,'C68');
xlswrite(outFile,0,1,'C69');
xlswrite(outFile,0,1,'C70');

xlswrite(outFile,PX,1,'C75');
xlswrite(outFile,PY,1,'C76');
xlswrite(outFile,PZ,1,'C77');
xlswrite(outFile,MX,1,'C78');
xlswrite(outFile,MY,1,'C79');
xlswrite(outFile,MZ,1,'C80');

xlswrite(outFile,sig,1,'C85:F85');
xlswrite(outFile,M,1,'C86:F86');

xlswrite(outFile,she1,1,'C91');
xlswrite(outFile,she2,1,'D91');
xlswrite(outFile,she3,1,'E91');
xlswrite(outFile,she4,1,'F91');
xlswrite(outFile,MS,1,'C92:F92');

xlswrite(outFile,tipw,1,'C97');
xlswrite(outFile,tipv,1,'C98');
xlswrite(outFile,twist,1,'C99');
