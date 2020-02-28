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
% Defines the inputs and output files
inFile = 'SE160A_2_AirLoads_Input.xlsx';
outFile= 'SE160A_2_AirLoads_Output.xlsx';

% Puts my name and student ID on the file
xlswrite(outFile, {'Steven Landgraf'},1,'C4');
xlswrite(outFile, {'A10850070'},1,'C5');

%Reading data from the input file and making the echo
[~,titleoutput]=xlsread(inFile,1,'C4');
xlswrite(outFile,titleoutput,1,'C7');

%Aircraft Geometry Definition Read
Read=xlsread(inFile,1,'C9:C13');
c4=Read(1); S=Read(2); b=Read(3); Cr=Read(4); Ct=Read(5);
hc4=xlsread(inFile,1,'C15');
xlswrite(outFile,Read,1,'C14:C18');
xlswrite(outFile,hc4,1,'C20');

%Aircraft Aerodynamic Definition
Read=xlsread(inFile,1,'C20:C31');
dclda=Read(1); a0=Read(2); clmaxp=Read(3); astallp=Read(4); clmaxn=Read(5);
astalln=Read(6); cmo=Read(7); cd0a=Read(8); cd0w=Read(9); cdiw=Read(10); 
d0=Read(11); d10=Read(12);
xlswrite(outFile,Read,1,'C25:C36');

%Aircraft Performance (at Sea Level)
%Flaps-Stowed
%Read=xlsread(inFile,1,'C37:C46');
Read=xlsread(inFile,1,'C37:C42');
vs1=Read(1); vno=Read(2); vne=Read(3); np=Read(4); nn=Read(5); %gustup0vc=Read(6);
%gustupvcvd=Read(7); gustd0vc=Read(8); gustdvcvd=Read(9); kg=Read(10);
kg=Read(6);
%xlswrite(outFile,Read,1,'C42:C51');
xlswrite(outFile,Read,1,'C42:C47');

%Flaps-Extended
%Read=xlsread(inFile,1,'C50:C53');
Read=xlsread(inFile,1,'C46:C49');
vs0=Read(1); vfe=Read(2); npf=Read(3); nnf=Read(4);
%xlswrite(outFile,Read,1,'C55:C58');
xlswrite(outFile,Read,1,'C51:C54');

%Aircraft Weight and Balance
%Read=xlsread(inFile,1,'C59:D65');
Read=xlsread(inFile,1,'C54:D60');
aircraft=Read(1,:); fuel=Read(2,:); pilot=Read(3,:); copilot=Read(4,:); 
passenger=Read(5,:); luggagepit=Read(6,:); luggagefuse=Read(7,:);
% xlswrite(outFile,Read,1,'C64:D70');
% xlswrite(outFile,Read,1,'C86:D92');
xlswrite(outFile,Read,1,'C59:D65');


%Aircraft Analysis Study
%Read=xlsread(inFile,1,'C70:C73');
Read=xlsread(inFile,1,'C65:C68');
h=Read(1); T=Read(2); nu=Read(3); WS=Read(4);
% xlswrite(outFile,Read,1,'C75:C78');
 xlswrite(outFile,Read,1,'C70:C73');

%Weight and Centroid Calculations
%Weight for each config
W(1)=aircraft(1)+pilot(1);
W(2)=W(1)+luggagepit(1);
W(3)=W(2)+copilot(1)+luggagefuse(1);
W(4)=W(3)+passenger(1);
W(5)=W(1)+fuel(1);
W(6)=W(5)+luggagepit(1);
W(7)=W(6)+copilot(1)+luggagefuse(1);
W(8)=W(7)+passenger(1);
%xlswrite(outFile,W,1,'E94:L94');
xlswrite(outFile,W,1,'E89:L89');

%Centroid for each component
x(1)=aircraft(1)*aircraft(2);
x(2)=fuel(1)*fuel(2);
x(3)=pilot(1)*pilot(2);
x(4)=copilot(1)*copilot(2);
x(5)=passenger(1)*passenger(2);
x(6)=luggagepit(1)*luggagepit(2);
x(7)=luggagefuse(1)*luggagefuse(2);

%Total centroid for each config
xc(1)=[x(1)+x(3)]/W(1);
xc(2)=[x(1)+x(3)+x(6)]/W(2);
xc(3)=[x(1)+x(3)+x(4)+x(6)+x(7)]/W(3);
xc(4)=[x(1)+x(3)+x(4)+x(5)+x(6)+x(7)]/W(4);
xc(5)=[x(1)+x(2)+x(3)]/W(5);
xc(6)=[x(1)+x(2)+x(3)+x(6)]/W(6);
xc(7)=[x(1)+x(2)+x(3)+x(4)+x(6)+x(7)]/W(7);
xc(8)=[x(1)+x(2)+x(3)+x(4)+x(5)+x(6)+x(7)]/W(8);
% xlswrite(outFile,xc,1,'E95:L95');
xlswrite(outFile,xc,1,'E90:L90');

hold on
plot(xc,W,'*')
xlabel('Xcg location (inch)')
ylabel('Weight(lb)')
grid on
hold off
figure

%Air Properties
Tempas=518.67-0.00356616*h;
Pas=2116.22*(Tempas/518.67)^5.255912;
rohair=Pas/((1716.5488)*(T+459.67));
Malt=(1.4*1716.5488*(T+459.67))^0.5;
rohsl=0.00237691;
sqr=(rohsl/rohair)^0.5;
%xlswrite(outFile,rohair,1,'C117');
xlswrite(outFile,rohair,1,'C112');

%V-n Diagram and Gust Diamond at maximum weight
V(1)=vs0;
V(2)=vs1;
V(3)=vno;
V(4)=vne;
V(5)=vne*1.2;
V(6)=vs1*np^0.5;
V(7)=vne;
V(8)=vs1*(-nn)^0.5;
V(9)=vne;
V(10)=vno;
V(11)=vne;
V(12)=vno;
V(13)=vne;
valt(1)=vs0*sqr;
valt(2)=vs1*sqr;
valt(3)=vno*sqr;
valt(4)=vne*sqr;
valt(5)=valt(4)*1.2;
valt(6)=valt(2)*np^0.5;
valt(7)=valt(4);
valt(8)=valt(2)*(-nn)^0.5;
valt(9)=valt(4);
valt(10)=valt(3);
valt(11)=valt(4);
valt(12)=valt(3);
valt(13)=valt(4);
for i=1:5
    n(i)=1;
    nalt(i)=1;
end
for i=6:7
    n(i)=np;
    nalt(i)=np;
end
for i=8:9
    n(i)=nn;
    nalt(i)=nn;
end
n(10)=1+kg*0.5*rohsl*50*(S/W(8))*dclda*(360/(2*pi))*(vno*(5280/3600));
nalt(10)=1+kg*0.5*rohair*50*(S/W(8))*dclda*(360/(2*pi))*(valt(3)*(5280/3600));
n(11)=1+kg*0.5*rohsl*25*(S/W(8))*dclda*(360/(2*pi))*(vne*(5280/3600));
nalt(11)=1+kg*0.5*rohair*25*(S/W(8))*dclda*(360/(2*pi))*(valt(4)*(5280/3600));
n(12)=1-kg*0.5*rohsl*50*(S/W(8))*dclda*(360/(2*pi))*(vno*(5280/3600));
nalt(12)=1-kg*0.5*rohair*50*(S/W(8))*dclda*(360/(2*pi))*(valt(3)*(5280/3600));
n(13)=1-kg*0.5*rohsl*25*(S/W(8))*dclda*(360/(2*pi))*(vne*(5280/3600));
nalt(13)=1-kg*0.5*rohair*25*(S/W(8))*dclda*(360/(2*pi))*(valt(4)*(5280/3600));
% xlswrite(outFile,V.',1,'C121:C133');
% xlswrite(outFile,n.',1,'D121:D133');
% xlswrite(outFile,valt.',1,'E121:E133');
% xlswrite(outFile,nalt.',1,'F121:F133');
 xlswrite(outFile,V.',1,'C116:C128');
 xlswrite(outFile,n.',1,'D116:D128');
 xlswrite(outFile,valt.',1,'E116:E128');
 xlswrite(outFile,nalt.',1,'F116:F128');


gustplot=[0 1; V(10) n(10); V(11) n(11); V(13) n(13); V(12) n(12); 0 1];
k=1;
zed=1;
for in=0:V(6)
    ngraph(k)=(in/vs1)^2;
    vplot(k)=in;
    k=k+1;
end
for in=V(6):V(7)
    ngraph(k)=np;
    vplot(k)=in;
    k=k+1;
end
for in=0:V(8)
    stuff(zed)=-(in/vs1)^2;
    rvplot(zed)=in;
    zed=zed+1;
end
    stuff(zed)=nn;
    rvplot(zed)=V(7);
loops=zed;
for efeofjse=0:loops-1
    ngraph(k)=stuff(zed);
    vplot(k)=rvplot(zed);
    k=k+1;
    zed=zed-1;
end
k=1;
for till=0:np
    flaps(k)=vs0*(till)^0.5;
    loadsgraphs(k)=till;
    k=k+1;
end
k=1;
for till=0:-nn
    flaps2(k)=vs0*(till)^0.5;
    loadsgraphs2(k)=-till;
    k=k+1;
end
k=1;
for ooops=nn:np
    flutter(k)=V(5);
    graphing(k)=ooops;
    k=k+1;
end
k=1;
for oops=0:V(13)
    horz(k)=1;
    horz2(k)=oops;
    k=k+1;
end
kool=[0 1; V(11) n(11)];
koolbeans=[0 1; V(13) n(13)];
hold on
plot(horz2,horz,'k')
plot(flutter,graphing,'k')
plot(flaps,loadsgraphs,'--b')
plot(flaps2,loadsgraphs2,'--b')
plot(gustplot(:,1),gustplot(:,2),'r')
plot(vplot,ngraph,'b')
plot(kool(:,1),kool(:,2),'--r')
plot(koolbeans(:,1),koolbeans(:,2),'--r')
grid on
ylabel('Load Factor(n)')
xlabel('Aircraft Speed - Sea Level (mph)')
hold off



%Part 3: Maneuvering Loads at Normal Flight (n=1)
dynamic(1)=0.5*rohair*(valt(1)*(5280/3600))^2;
dynamic(2)=0.5*rohair*(valt(2)*(5280/3600))^2;
dynamic(3)=0.5*rohair*(valt(3)*(5280/3600))^2;
dynamic(4)=0.5*rohair*(valt(4)*(5280/3600))^2;
loadfact(1)=1;
loadfact(2)=1;
loadfact(3)=1;
loadfact(4)=1;
wingmoment(1)=dynamic(1)*cmo*S*S/b;
wingmoment(2)=dynamic(2)*cmo*S*S/b;
wingmoment(3)=dynamic(3)*cmo*S*S/b;
wingmoment(4)=dynamic(4)*cmo*S*S/b;
winglift(1)=12*(((c4-xc(8))/12)*-W(8)+wingmoment(1))/(hc4-c4);
winglift(2)=12*(((c4-xc(8))/12)*-W(8)+wingmoment(2))/(hc4-c4);
winglift(3)=12*(((c4-xc(8))/12)*-W(8)+wingmoment(3))/(hc4-c4);
winglift(4)=12*(((c4-xc(8))/12)*-W(8)+wingmoment(4))/(hc4-c4);
lift(1)=W(8)-winglift(1);
lift(2)=W(8)-winglift(2);
lift(3)=W(8)-winglift(3);
lift(4)=W(8)-winglift(4);
CL(1)=lift(1)/(dynamic(1)*S);
CL(2)=lift(2)/(dynamic(2)*S);
CL(3)=lift(3)/(dynamic(3)*S);
CL(4)=lift(4)/(dynamic(4)*S);
walpha(1)=(CL(1)/dclda)+a0;
walpha(2)=(CL(2)/dclda)+a0;
walpha(3)=(CL(3)/dclda)+a0;
walpha(4)=(CL(4)/dclda)+a0;
fusedrag(1)=S*cd0a*dynamic(1);
fusedrag(2)=S*cd0a*dynamic(2);
fusedrag(3)=S*cd0a*dynamic(3);
fusedrag(4)=S*cd0a*dynamic(4);
wingdrag(1)=dynamic(1)*S*(cd0w+cdiw*CL(1)^2);
wingdrag(2)=dynamic(2)*S*(cd0w+cdiw*CL(2)^2);
wingdrag(3)=dynamic(3)*S*(cd0w+cdiw*CL(3)^2);
wingdrag(4)=dynamic(4)*S*(cd0w+cdiw*CL(4)^2);
thrust(1)=fusedrag(1)+wingdrag(1);
thrust(2)=fusedrag(2)+wingdrag(2);
thrust(3)=fusedrag(3)+wingdrag(3);
thrust(4)=fusedrag(4)+wingdrag(4);
althp(1)=valt(1)*(5280/3600)*thrust(1)/(nu/100)/550;
althp(2)=valt(2)*(5280/3600)*thrust(2)/(nu/100)/550;
althp(3)=valt(3)*(5280/3600)*thrust(3)/(nu/100)/550;
althp(4)=valt(4)*(5280/3600)*thrust(4)/(nu/100)/550;
slhp(1)=althp(1)*((rohair/rohsl)-(1-(rohair/rohsl))/7.55)^-1;
slhp(2)=althp(2)*((rohair/rohsl)-(1-(rohair/rohsl))/7.55)^-1;
slhp(3)=althp(3)*((rohair/rohsl)-(1-(rohair/rohsl))/7.55)^-1;
slhp(4)=althp(4)*((rohair/rohsl)-(1-(rohair/rohsl))/7.55)^-1;
 
 xlswrite(outFile,valt(1:4),1,'C157:F157');
 xlswrite(outFile,dynamic(1:4),1,'C158:F158');
 xlswrite(outFile,loadfact(1:4),1,'C159:F159');
 xlswrite(outFile,wingmoment(1:4),1,'C160:F160');
 xlswrite(outFile,lift(1:4),1,'C161:F161');
 xlswrite(outFile,CL(1:4),1,'C162:F162');
 xlswrite(outFile,winglift(1:4),1,'C163:F163');
 xlswrite(outFile,wingdrag(1:4),1,'C164:F164');
 xlswrite(outFile,fusedrag(1:4),1,'C165:F165');
 xlswrite(outFile,thrust(1:4),1,'C166:F166');
 xlswrite(outFile,althp(1:4),1,'C167:F167');
 xlswrite(outFile,slhp(1:4),1,'C168:F168');
 xlswrite(outFile,walpha(1:4),1,'C169:F169');

%Part 4: Maneuvering Loads at Critical V-n Diagram Corners and Gust Diamond
%Corners
dynamic(5)=0.5*rohsl*(V(6)*(5280/3600))^2;
dynamic(6)=0.5*rohsl*(V(7)*(5280/3600))^2;
dynamic(7)=0.5*rohsl*(V(8)*(5280/3600))^2;
dynamic(8)=0.5*rohsl*(V(9)*(5280/3600))^2;
dynamic(9)=0.5*rohsl*(V(10)*(5280/3600))^2;
dynamic(10)=0.5*rohsl*(V(11)*(5280/3600))^2;
dynamic(11)=0.5*rohsl*(V(12)*(5280/3600))^2;
dynamic(12)=0.5*rohsl*(V(13)*(5280/3600))^2;
loadfact(5)=n(6);
loadfact(6)=n(7);
loadfact(7)=n(8);
loadfact(8)=n(9);
loadfact(9)=n(10);
loadfact(10)=n(11);
loadfact(11)=n(12);
loadfact(12)=n(13);
wingmoment(5)=dynamic(5)*cmo*S*S/b;
wingmoment(6)=dynamic(6)*cmo*S*S/b;
wingmoment(7)=dynamic(7)*cmo*S*S/b;
wingmoment(8)=dynamic(8)*cmo*S*S/b;
wingmoment(9)=dynamic(9)*cmo*S*S/b;
wingmoment(10)=dynamic(10)*cmo*S*S/b;
wingmoment(11)=dynamic(11)*cmo*S*S/b;
wingmoment(12)=dynamic(12)*cmo*S*S/b;
winglift(5)=12*(((c4-xc(8))/12)*-loadfact(5)*W(8)+wingmoment(5))/(hc4-c4);
winglift(6)=12*(((c4-xc(8))/12)*-loadfact(6)*W(8)+wingmoment(6))/(hc4-c4);
winglift(7)=12*(((c4-xc(8))/12)*-loadfact(7)*W(8)+wingmoment(7))/(hc4-c4);
winglift(8)=12*(((c4-xc(8))/12)*-loadfact(8)*W(8)+wingmoment(8))/(hc4-c4);
winglift(9)=12*(((c4-xc(8))/12)*-loadfact(9)*W(8)+wingmoment(9))/(hc4-c4);
winglift(10)=12*(((c4-xc(8))/12)*-loadfact(10)*W(8)+wingmoment(10))/(hc4-c4);
winglift(11)=12*(((c4-xc(8))/12)*-loadfact(11)*W(8)+wingmoment(11))/(hc4-c4);
winglift(12)=12*(((c4-xc(8))/12)*-loadfact(12)*W(8)+wingmoment(12))/(hc4-c4);
lift(5)=loadfact(5)*W(8)-winglift(5);
lift(6)=loadfact(6)*W(8)-winglift(6);
lift(7)=loadfact(7)*W(8)-winglift(7);
lift(8)=loadfact(8)*W(8)-winglift(8);
lift(9)=loadfact(9)*W(8)-winglift(9);
lift(10)=loadfact(10)*W(8)-winglift(10);
lift(11)=loadfact(11)*W(8)-winglift(11);
lift(12)=loadfact(12)*W(8)-winglift(12);
CL(5)=lift(5)/(dynamic(5)*S);
CL(6)=lift(6)/(dynamic(6)*S);
CL(7)=lift(7)/(dynamic(7)*S);
CL(8)=lift(8)/(dynamic(8)*S);
CL(9)=lift(9)/(dynamic(9)*S);
CL(10)=lift(10)/(dynamic(10)*S);
CL(11)=lift(11)/(dynamic(11)*S);
CL(12)=lift(12)/(dynamic(12)*S);
walpha(5)=(CL(5)/dclda)+a0;
walpha(6)=(CL(6)/dclda)+a0;
walpha(7)=(CL(7)/dclda)+a0;
walpha(8)=(CL(8)/dclda)+a0;
walpha(9)=(CL(9)/dclda)+a0;
walpha(10)=(CL(10)/dclda)+a0;
walpha(11)=(CL(11)/dclda)+a0;
walpha(12)=(CL(12)/dclda)+a0;
fusedrag(5)=S*cd0a*dynamic(5);
fusedrag(6)=S*cd0a*dynamic(6);
fusedrag(7)=S*cd0a*dynamic(7);
fusedrag(8)=S*cd0a*dynamic(8);
fusedrag(9)=S*cd0a*dynamic(9);
fusedrag(10)=S*cd0a*dynamic(10);
fusedrag(11)=S*cd0a*dynamic(11);
fusedrag(12)=S*cd0a*dynamic(12);
wingdrag(5)=dynamic(5)*S*(cd0w+cdiw*CL(5)^2);
wingdrag(6)=dynamic(6)*S*(cd0w+cdiw*CL(6)^2);
wingdrag(7)=dynamic(7)*S*(cd0w+cdiw*CL(7)^2);
wingdrag(8)=dynamic(8)*S*(cd0w+cdiw*CL(8)^2);
wingdrag(9)=dynamic(9)*S*(cd0w+cdiw*CL(9)^2);
wingdrag(10)=dynamic(10)*S*(cd0w+cdiw*CL(10)^2);
wingdrag(11)=dynamic(11)*S*(cd0w+cdiw*CL(11)^2);
wingdrag(12)=dynamic(12)*S*(cd0w+cdiw*CL(12)^2);
thrust(5)=fusedrag(5)+wingdrag(5);
thrust(6)=fusedrag(6)+wingdrag(6);
thrust(7)=fusedrag(7)+wingdrag(7);
thrust(8)=fusedrag(8)+wingdrag(8);
thrust(9)=fusedrag(9)+wingdrag(9);
thrust(10)=fusedrag(10)+wingdrag(10);
thrust(11)=fusedrag(11)+wingdrag(11);
thrust(12)=fusedrag(12)+wingdrag(12);
slhp(5)=V(6)*(5280/3600)*thrust(5)/(nu/100)/550;
slhp(6)=V(7)*(5280/3600)*thrust(6)/(nu/100)/550;
slhp(7)=V(8)*(5280/3600)*thrust(7)/(nu/100)/550;
slhp(8)=V(9)*(5280/3600)*thrust(8)/(nu/100)/550;
slhp(9)=V(10)*(5280/3600)*thrust(9)/(nu/100)/550;
slhp(10)=V(11)*(5280/3600)*thrust(10)/(nu/100)/550;
slhp(11)=V(12)*(5280/3600)*thrust(11)/(nu/100)/550;
slhp(12)=V(13)*(5280/3600)*thrust(12)/(nu/100)/550;
 
xlswrite(outFile,V(6:13),1,'C178:J178');
 xlswrite(outFile,dynamic(5:12),1,'C179:J179');
 xlswrite(outFile,loadfact(5:12),1,'C180:J180');
 xlswrite(outFile,wingmoment(5:12),1,'C181:J181');
 xlswrite(outFile,lift(5:12),1,'C182:J182');
 xlswrite(outFile,CL(5:12),1,'C183:J183');
 xlswrite(outFile,winglift(5:12),1,'C184:J184');
 xlswrite(outFile,wingdrag(5:12),1,'C185:J185');
 xlswrite(outFile,fusedrag(5:12),1,'C186:J186');
 xlswrite(outFile,thrust(5:12),1,'C187:J187');
 xlswrite(outFile,slhp(5:12),1,'C188:J188');
 xlswrite(outFile,walpha(5:12),1,'C189:J189');

%Part 5 Distributed Loads at PHAA
k=1;
for i=0:0.1:(b/2)
    c(k)=Cr*(0.5-(i/b)*(1-Ct/Cr)+((1+Ct/Cr)/pi)*(1-((2*i)/b)^2)^0.5);
    dweight(k)=(loadfact(5)*W(8)*Cr)/S;
    l(k)=dynamic(5)*CL(5)*c(k);
    d(k)=dynamic(5)*Cr*(1-(2*i*(1-Ct/Cr)/b))*(cd0w+cdiw*CL(5)^2);
    m(k)=dynamic(5)*(Cr^2)*(1-((2*i)/b)*(1-Ct/Cr))^2*cmo;
    fn(k)=cosd(walpha(5))*(l(k)-loadfact(5)*W(8))+sind(walpha(5))*d(k);
    fc(k)=(loadfact(5)-l(k))*sind(walpha(5))+cosd(walpha(5))*d(k);
    base(k)=i;
    k=k+1;
end
figure
hold on
plot(base,dweight,'k')
plot(-base,dweight,'k')
grid on
xlabel('span (feet)')
ylabel('Dynamic Weight nWc/S (lb/ft)')
hold off

figure
hold on
plot(base,l,'k')
plot(-base,l,'k')
grid on
xlabel('span (feet)')
ylabel('Distributed Lift (lb/ft)')
hold off

figure
hold on
plot(base,d,'k')
plot(-base,d,'k')
grid on
xlabel('span (feet)')
ylabel('Distributed Drag (lb/ft)')
hold off

figure
hold on
plot(base,m,'k')
plot(-base,m,'k')
grid on
xlabel('span (feet)')
ylabel('Distributed Twist Moment mt(lb/ft)')
hold off

figure
hold on
plot(base,fn,'k')
plot(-base,fn,'k')
grid on
xlabel('span (feet)')
ylabel('Distributed Normal Force(lb/ft)')
hold off

figure
hold on
plot(base,fc,'k')
plot(-base,fc,'k')
grid on
xlabel('span (feet)')
ylabel('Distributed Chord Force(lb/ft)')
hold off

createFigure('SE160A_2_AirLoads_Output.xlsx', 1, 1, 'B92', 'J106');
createFigure('SE160A_2_AirLoads_Output.xlsx', 1, 2, 'B130', 'J150');
createFigure('SE160A_2_AirLoads_Output.xlsx', 1, 3, 'B195', 'J215');
createFigure('SE160A_2_AirLoads_Output.xlsx', 1, 4, 'B217', 'J237');
createFigure('SE160A_2_AirLoads_Output.xlsx', 1, 5, 'B239', 'J259');
createFigure('SE160A_2_AirLoads_Output.xlsx', 1, 6, 'B261', 'J281');
createFigure('SE160A_2_AirLoads_Output.xlsx', 1, 7, 'B283', 'J303');
createFigure('SE160A_2_AirLoads_Output.xlsx', 1, 8, 'B305', 'H325');

