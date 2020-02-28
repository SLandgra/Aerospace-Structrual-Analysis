% . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
% .
% .  SE-160A / 260A: Mechanics of Aerospace Structures
% .
% .  Title:   MATLAB Refresher on Strength of Materials and Linear Algebra
% .  Author:  Steven Landgraf                    
% .  PID:     A10850070
% .  Revised: 
% .  
% .  This program calculates the section properties of an arbitrary cross-
% .  section that is comprised of thin-wall straight "skin" segments
% .  and/or discrete "stringer" elements. This program takes in (x,y)
% .  coordinates for the skin segments, Area and inertial properties for
% .  the stringers, and a possible user generated origin. This program 
% .  outputs the inertial properties of the figure about the origin,
% .  centroid, and the possible inputted origin.
% . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
clc
clear all
close all
format compact

% Defines the inputs and output files
inFile = 'SE160A_1_Section_Input.xlsx';
%inFile = 'SE160A_1_Section_Input.xlsx';
outFile= 'SE160A_1_Section_Output.xlsx';

% Puts my name and student ID on the file
xlswrite(outFile, {'Steven Landgraf'},1,'C4');
xlswrite(outFile, {'A10850070'},1,'C5');

%Reading data from the input file and making the echo
[~,titleoutput]=xlsread(inFile,1,'C4');
xlswrite(outFile,titleoutput,1,'C7');

n1=xlsread(inFile,1,'B8');
xlswrite(outFile,n1,1,'B13');

x11=xlsread(inFile,1,'B11:B30');
xlswrite(outFile,x11,1,'B16:B35');

y11=xlsread(inFile,1,'C11:C30');
xlswrite(outFile,y11,1,'C16:C35');

x12=xlsread(inFile,1,'D11:D30');
xlswrite(outFile,x12,1,'D16:D35');

y12=xlsread(inFile,1,'E11:E30');
xlswrite(outFile,y12,1,'E16:E35');

t1=xlsread(inFile,1,'F11:F30');
xlswrite(outFile,t1,1,'F16:F35');

n2=xlsread(inFile,1,'B34');
if n2==0                                        %Conditional to check if no spars & will ignore the rest of the spar portions if no spars
xlswrite(outFile,0,1,'B39');
else
xlswrite(outFile,n2,1,'B39');

x21=xlsread(inFile,1,'B37:B41');
xlswrite(outFile,x21,1,'B42:B46');

y21=xlsread(inFile,1,'C37:C41');
xlswrite(outFile,y21,1,'C42:C46');

x22=xlsread(inFile,1,'D37:D41');
xlswrite(outFile,x22,1,'D42:D46');

y22=xlsread(inFile,1,'E37:E41');
xlswrite(outFile,y22,1,'E42:E46');

t2=xlsread(inFile,1,'F37:F41');
xlswrite(outFile,t2,1,'F42:F46');
end

n3=xlsread(inFile,1,'B45');
if n3==0                                        %Conditional to check if no stringers & will ignore the rest of the stringer portions if no spars
xlswrite(outFile,0,1,'B50');
else
xlswrite(outFile,n3,1,'B50');

x3=xlsread(inFile,1,'B48:B57');
xlswrite(outFile,x3,1,'B53:B62');

y3=xlsread(inFile,1,'C48:C57');
xlswrite(outFile,y3,1,'C53:C62');

as=xlsread(inFile,1,'D48:D57');
xlswrite(outFile,as,1,'D53:D62');

ixxs=xlsread(inFile,1,'E48:E57');
xlswrite(outFile,ixxs,1,'E53:E62');

iyys=xlsread(inFile,1,'F48:F57');
xlswrite(outFile,iyys,1,'F53:F62');

ixys=xlsread(inFile,1,'G48:G57');
xlswrite(outFile,ixys,1,'G53:G62');
end

ox=xlsread(inFile,1,'B62');
xlswrite(outFile,ox,1,'B67');

oy=xlsread(inFile,1,'C62');
xlswrite(outFile,oy,1,'C67');

%Initializations for each loop
SkinCentroidx=0;
SkinCentroidy=0;
Ixxskin=0;
Iyyskin=0;
Ixyskin=0;
Askintotal=0;
%Skin Calculations
for i=1:n1
    L=((x12(i)-x11(i))^2+(y12(i)-y11(i))^2)^0.5;                                            %Calculates the length of each segment for the skin
    Askin=t1(i)*L;                                                                          %Area of each skin segment
    SkinCentroidx= (0.5*Askin*(x11(i)+x12(i)))+SkinCentroidx;                               %Summation of the skin portion for the x centroid
    SkinCentroidy= (0.5*Askin*(y11(i)+y12(i)))+SkinCentroidy;                               %Summation of the skin portion for the y centroid
    Ixxskin=(Askin/6)*(y11(i)*(2*y11(i)+y12(i))+y12(i)*(y11(i)+2*y12(i)))+Ixxskin;          %the Ixx summation portion for the skin
    Iyyskin=(Askin/6)*(x11(i)*(2*x11(i)+x12(i))+x12(i)*(x11(i)+2*x12(i)))+Iyyskin;          %the Iyy summation portion for the skin
    Ixyskin=(Askin/6)*((x11(i)*(2*y11(i)+y12(i)))+(x12(i)*(y11(i)+2*y12(i))))+Ixyskin;      %the Ixy summation portion for the skin
    Askintotal=Askintotal+Askin;                                                            %Summation of the skin area
end
%Same loop as the skin loop for calculations, but for the spars. Will not
%work if there are no spars
if n2~=0
for i=1:n2
    L=((x22(i)-x21(i))^2+(y22(i)-y21(i))^2)^0.5;
    Askin=t2(i)*L;
    SkinCentroidx= (0.5*Askin*(x21(i)+x22(i)))+SkinCentroidx;
    SkinCentroidy= (0.5*Askin*(y21(i)+y22(i)))+SkinCentroidy;
    Ixxskin=(Askin/6)*(y21(i)*(2*y21(i)+y22(i))+y22(i)*(y21(i)+2*y22(i)))+Ixxskin;
    Iyyskin=(Askin/6)*(x21(i)*(2*x21(i)+x22(i))+x22(i)*(x21(i)+2*x22(i)))+Iyyskin;
    Ixyskin=(Askin/6)*((x21(i)*(2*y21(i)+y22(i)))+(x22(i)*(y21(i)+2*y22(i))))+Ixyskin;
    Askintotal=Askintotal+Askin;
end
end
%Initializations for stringer loop
StringCentroidx=0;
StringCentroidy=0;
Astringtotal=0;
Ixxstring=0;
Ixystring=0;
Iyystring=0;
%Stringer summation calculations will not work if no stringers
if n3~=0
for i=1:n3
    StringCentroidx= (x3(i)*as(i))+StringCentroidx;                                         %Stringer x centroid summation portion
    StringCentroidy= (y3(i)*as(i))+StringCentroidy;                                         %Stringer y centroid summation portion
    Astringtotal=as(i)+Astringtotal;                                                        %Total area of the stringers
    Ixxstring=ixxs(i)+as(i)*(y3(i))^2+Ixxstring;                                            %Ixx summation portion for stringer
    Iyystring=iyys(i)+as(i)*(x3(i))^2+Iyystring;                                            %Iyy summation portion for stringer
    Ixystring=ixys(i)+(as(i)*(y3(i)*x3(i)))+Ixystring;                                      %Ixy summation portion for stringer
end
end
%Total area and the write portion to the output file
A=Askintotal+Astringtotal;
xlswrite(outFile,A,1,'D75');
xlswrite(outFile,A,1,'D149');

%Total x Centroid and the write portion to the output file
xCentroid=(SkinCentroidx+StringCentroidx)/A;
xlswrite(outFile,xCentroid,1,'D76');

%Total y Centroid and the write portion to the output file
yCentroid=(SkinCentroidy+StringCentroidy)/A;
xlswrite(outFile,yCentroid,1,'D77');

%Total Ixx and the write portion to the output file
Ixx=Ixxskin+Ixxstring;
xlswrite(outFile,Ixx,1,'D78');

%Total Iyy and the write portion to the output file
Iyy=Iyyskin+Iyystring;
xlswrite(outFile,Iyy,1,'D79');

%Total Ixy and the write portion to the output file
Ixy=Ixyskin+Ixystring;
xlswrite(outFile,Ixy,1,'D80');

%Calculation for Ixxcentroid and the write portion to the output file
Ixxcentroid=Ixx-A*yCentroid^2;
xlswrite(outFile,Ixxcentroid,1,'D111');

%Calculation for Iyycentroid and the write portion to the output file
Iyycentroid=Iyy-A*xCentroid^2;
xlswrite(outFile,Iyycentroid,1,'D112');

%Calculation for Ixycentroid and the write portion to the output file
Ixycentroid=Ixy-A*xCentroid*yCentroid;
xlswrite(outFile,Ixycentroid,1,'D113');

%Alpha calculation and write portion
alpha=0.5*atand(-(2*Ixycentroid)/(Ixxcentroid-Iyycentroid));
xlswrite(outFile,alpha,1,'D115');

%Calculation for I11 and the write portion to the output file
I11=Ixxcentroid*(cosd(alpha)^2)+Iyycentroid*(sind(alpha)^2)-2*Ixycentroid*cosd(alpha)*sind(alpha);
xlswrite(outFile,I11,1,'D116');

%Calculation for I22 and the write portion to the output file
I22=Ixxcentroid*(sind(alpha)^2)+Iyycentroid*(cosd(alpha)^2)+2*Ixycentroid*cosd(alpha)*sind(alpha);
xlswrite(outFile,I22,1,'D117');

%Calculation for I12 and the write portion to the output file
I12=(Ixxcentroid-Iyycentroid)*cosd(alpha)*sind(alpha)+Ixycentroid*((cosd(alpha)^2)-(sind(alpha)^2));
xlswrite(outFile,I12,1,'D118');

%Write for x Centroid
xlswrite(outFile,xCentroid-ox,1,'D150');

%Write for y Centroid
xlswrite(outFile,yCentroid-oy,1,'D151');

%Ixx at the user origin calculation and write
Ixxorgin=Ixxcentroid+A*(oy-yCentroid)^2;
xlswrite(outFile,Ixxorgin,1,'D152');

%Iyy at the user origin calculation and write
Iyyorgin=Iyycentroid+A*(ox-xCentroid)^2;
xlswrite(outFile,Iyyorgin,1,'D153');

%Ixy at the user origin calculation and write
Ixyorgin=Ixycentroid+A*(oy-yCentroid)*(ox-xCentroid);
xlswrite(outFile,Ixyorgin,1,'D154');

%Plotting section
hold on
%Skin plot
for i=1:n1
    k=[x11(i) x12(i)];  %These arrays make it easier to plot each skin segment
    h=[y11(i) y12(i)];
    plot(k,h,'b')
end
%Spar plot (same as skin)
for i=1:n2
    k=[x21(i) x22(i)];
    h=[y21(i) y22(i)];
    plot(k,h,'b')
end
%Stringer plot
for i=1:n3
    scatter(x3(i),y3(i),'b','*') %does the stringers as a scatter plot
end
axis('equal') %makes the axis nice to look at
title('Initial Reference Frame')
grid on
ax=gca;
ax.XAxisLocation='origin'; %puts the orgin in the middle of the figure
ax.YAxisLocation='origin';
hold off

figure %figure 2 for the next graph each loop corresponds like the last one
hold on
for i=1:n1
    k=[x11(i)-xCentroid x12(i)-xCentroid];
    h=[y11(i)-yCentroid y12(i)-yCentroid];
    plot(k,h,'b')
end
for i=1:n2
    k=[x21(i)-xCentroid x22(i)-xCentroid];
    h=[y21(i)-yCentroid y22(i)-yCentroid];
    plot(k,h,'b')
end
for i=1:n3
    scatter(x3(i)-xCentroid,y3(i)-yCentroid,'b','*')
end
po=axis;
ew=1;
%Adds in the rotated axis (x portion)
for i=0:po(2)
    axisxx(ew)=i*cosd(-alpha); %rotates the axis and gets the x values for the rotation
    axisxy(ew)=-i*sind(-alpha); %rotates the axis and gets the y values for the rotation
    ew=ew+1;  %helps with itteration
end
ew=1;
plot(axisxx,axisxy,'r');
%Adds in the rotated axis (y portion)
for i2=0:po(4)
     axisyx(ew)=i2*sind(-alpha);    
     axisyy(ew)=i2*cosd(-alpha);
     ew=ew+1;
end
plot(axisyx,axisyy,'r');
axis('equal')
title('Centroid Reference Frame')
grid on
ax=gca;
ax.XAxisLocation='origin';
ax.YAxisLocation='origin';
hold off
%Third plot for the user generated origin
figure
hold on
for i=1:n1
    k=[x11(i)-ox x12(i)-ox];
    h=[y11(i)-oy y12(i)-oy];
    plot(k,h,'b')
end
for i=1:n2
    k=[x21(i)-ox x22(i)-ox];
    h=[y21(i)-oy y22(i)-oy];
    plot(k,h,'b')
end
for i=1:n3
    scatter(x3(i)-ox,y3(i)-oy,'b','*')
end
axis('equal')
title('User-Defined Reference Frame')
grid on
ax=gca;
ax.XAxisLocation='origin';
ax.YAxisLocation='origin';
hold off
%puts the figures into the output file.
createFigure('SE160A_1_Section_Output.xlsx', 1, 1, 'B82', 'H106');
createFigure('SE160A_1_Section_Output.xlsx', 1, 2, 'B120', 'H144');
createFigure('SE160A_1_Section_Output.xlsx', 1, 3, 'B156', 'H180');