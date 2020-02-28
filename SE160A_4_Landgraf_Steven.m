close all
clear all
clc
format compact
inFile = 'SE160A_4_Metallics_Input.xlsx';
outFile = 'SE160A_4_Metallics_Output.xlsx';

xlswrite(outFile, {'Steven Landgraf'}, 1, 'C4');
xlswrite(outFile, {'A10850070'}  , 1, 'C5');
%Input Read and Echo
[~,titles] = xlsread(inFile,1, 'C5');
xlswrite(outFile,titles,1,'C7');

%Part 1 Read and Write
Cy=xlsread(inFile,1,'C11:D20');
Cu=xlsread(inFile,1,'E11:F20');
Ty=xlsread(inFile,1,'G11:H20');
Tu=xlsread(inFile,1,'I11:J20');
xlswrite(outFile,Cy,1,'C15:D24');
xlswrite(outFile,Cu,1,'E15:F24');
xlswrite(outFile,Ty,1,'G15:H24');
xlswrite(outFile,Tu,1,'I15:J24');

%Mean calculations for first 6
mCy=mean(Cy(1:6,:));
mCu=mean(Cu(1:6,:));
mTy=mean(Ty(1:6,:));
mTu=mean(Tu(1:6,:));
xlswrite(outFile,mCy,1,'C30:D30');
xlswrite(outFile,mCu,1,'E30:F30');
xlswrite(outFile,mTy,1,'G30:H30');
xlswrite(outFile,mTu,1,'I30:J30');

%Standard deviation equationsfor first 6
sCy=std(Cy(1:6,:));
sCu=std(Cu(1:6,:));
sTy=std(Ty(1:6,:));
sTu=std(Tu(1:6,:));
xlswrite(outFile,sCy,1,'C31:D31');
xlswrite(outFile,sCu,1,'E31:F31');
xlswrite(outFile,sTy,1,'G31:H31');
xlswrite(outFile,sTu,1,'I31:J31');

%K constants
Ka6=5.062;
Kb6=3.007;

%A basis calculations for first 6
AaCy=mCy+Ka6*sCy;
AaCu=mCu+Ka6*sCu;
AaTy=mTy-Ka6*sTy;
AaTu=mTu-Ka6*sTu;
xlswrite(outFile,AaCy,1,'C32:D32');
xlswrite(outFile,AaCu,1,'E32:F32');
xlswrite(outFile,AaTy,1,'G32:H32');
xlswrite(outFile,AaTu,1,'I32:J32');

%B basis calculations for first 6
AbCy=mCy+Kb6*sCy;
AbCu=mCu+Kb6*sCu;
AbTy=mTy-Kb6*sTy;
AbTu=mTu-Kb6*sTu;
xlswrite(outFile,AbCy,1,'C33:D33');
xlswrite(outFile,AbCu,1,'E33:F33');
xlswrite(outFile,AbTy,1,'G33:H33');
xlswrite(outFile,AbTu,1,'I33:J33');

%S basis calculations for first 6
AsCy=max(Cy(1:6,:));
AsCu=max(Cu(1:6,:));
AsTy=min(Ty(1:6,:));
AsTu=min(Tu(1:6,:));
xlswrite(outFile,AsCy,1,'C34:D34');
xlswrite(outFile,AsCu,1,'E34:F34');
xlswrite(outFile,AsTy,1,'G34:H34');
xlswrite(outFile,AsTu,1,'I34:J34');

%Mean calculations for all 10
mCy=mean(Cy);
mCu=mean(Cu);
mTy=mean(Ty);
mTu=mean(Tu);
xlswrite(outFile,mCy,1,'C38:D38');
xlswrite(outFile,mCu,1,'E38:F38');
xlswrite(outFile,mTy,1,'G38:H38');
xlswrite(outFile,mTu,1,'I38:J38');

%Standard devation calculations for all 10
sCy=std(Cy);
sCu=std(Cu);
sTy=std(Ty);
sTu=std(Tu);
xlswrite(outFile,sCy,1,'C39:D39');
xlswrite(outFile,sCu,1,'E39:F39');
xlswrite(outFile,sTy,1,'G39:H39');
xlswrite(outFile,sTu,1,'I39:J39');

%K constants
Ka10=3.981;
Kb10=2.355;
Ks10=1.5099;

%A basis calculations for all 10
AaCy1=mCy+Ka10*sCy;
AaCu1=mCu+Ka10*sCu;
AaTy1=mTy-Ka10*sTy;
AaTu1=mTu-Ka10*sTu;
xlswrite(outFile,AaCy1,1,'C40:D40');
xlswrite(outFile,AaCu1,1,'E40:F40');
xlswrite(outFile,AaTy1,1,'G40:H40');
xlswrite(outFile,AaTu1,1,'I40:J40');

%B basis calculations for all 10
AbCy1=mCy+Kb10*sCy;
AbCu1=mCu+Kb10*sCu;
AbTy1=mTy-Kb10*sTy;
AbTu1=mTu-Kb10*sTu;
xlswrite(outFile,AbCy1,1,'C41:D41');
xlswrite(outFile,AbCu1,1,'E41:F41');
xlswrite(outFile,AbTy1,1,'G41:H41');
xlswrite(outFile,AbTu1,1,'I41:J41');

%S basis calculations for all 10
AsCy1=max(Cy);
AsCu1=max(Cu);
AsTy1=min(Ty);
AsTu1=min(Tu);
xlswrite(outFile,AsCy1,1,'C42:D42');
xlswrite(outFile,AsCu1,1,'E42:F42');
xlswrite(outFile,AsTy1,1,'G42:H42');
xlswrite(outFile,AsTu1,1,'I42:J42');

%Part 2 read and write
sig=xlsread(inFile,1,'C26:C28');
she=xlsread(inFile,1,'C29:C31');
fact=xlsread(inFile,1,'C32:C33');
xlswrite(outFile,sig,1,'C50:C52');
xlswrite(outFile,she,1,'C53:C55');
xlswrite(outFile,fact,1,'C56:C57');

%Allowable stresses tension and compression A basis
sigtstarA=min([AaTy1(1)/fact(1) AaTu1(1)/fact(2)])*1000;
sigcstarA=max([AaCy1(1)/fact(1) AaCu1(1)/fact(2)])*1000;
xlswrite(outFile,sigtstarA,1,'C63');
xlswrite(outFile,sigcstarA,1,'C64');

%Allowable strains tension and compression A basis
epststarA=min([AaTy1(2)/fact(1) AaTu1(2)/fact(2)])*1000;
epscstarA=max([AaCy1(2)/fact(1) AaCu1(2)/fact(2)])*1000;
xlswrite(outFile,epststarA,1,'C70');
xlswrite(outFile,epscstarA,1,'C71');

%Shear for stress and strains for tresca A basis
if sigtstarA~=0&&sigcstarA~=0
sheartressigA=0.25*(sigtstarA-sigcstarA);
sheartresepsA=0.25*(epststarA-epscstarA);
else
    sheartressigA=0.5*(sigtstarA+sigcstarA);
    sheartresepsA=0.5*(epststarA+epscstarA);
end
xlswrite(outFile,sheartressigA,1,'C66');
xlswrite(outFile,sheartresepsA,1,'C73');

%Shear for stress and strains for Rankine A basis
shearranksigA=(sigtstarA-sigcstarA)/2;
shearranksepsA=(epststarA-epscstarA)/2;
xlswrite(outFile,shearranksigA,1,'C65');
xlswrite(outFile,shearranksepsA,1,'C72');

%Allowable stresses tension and compression B basis
sigtstarB=min([AbTy1(1)/fact(1) AbTu1(1)/fact(2)])*1000;
sigcstarB=max([AbCy1(1)/fact(1) AbCu1(1)/fact(2)])*1000;
xlswrite(outFile,sigtstarB,1,'D63');
xlswrite(outFile,sigcstarB,1,'D64');

%Allowable strains tension and compression B basis
epststarB=min([AbTy1(2)/fact(1) AbTu1(2)/fact(2)])*1000;
epscstarB=max([AbCy1(2)/fact(1) AbCu1(2)/fact(2)])*1000;
xlswrite(outFile,epststarB,1,'D70');
xlswrite(outFile,epscstarB,1,'D71');

%Shear for stress and strains for tresca B basis
if sigtstarB~=0&&sigcstarB~=0
sheartressigB=0.25*(sigtstarB-sigcstarB);
sheartresepsB=0.25*(epststarB-epscstarB);
else
    sheartressigB=0.5*(sigtstarB+sigcstarB);
    sheartresepsB=0.5*(epststarB+epscstarB);
end
xlswrite(outFile,sheartressigB,1,'D66')
xlswrite(outFile,sheartresepsB,1,'D73')

%Shear for stress and strains for Rankine A basis
shearranksigB=(sigtstarB-sigcstarB)/2;
shearrankepsB=(epststarB-epscstarB)/2;
xlswrite(outFile,shearranksigB,1,'D65');
xlswrite(outFile,shearrankepsB,1,'D72');

%Principle Stresses calculations w/ tmax
A=[sig(1) she(3) she(2);
    she(3) sig(2) she(1);
    she(2) she(1) sig(3)];
B=eig(A);
if B(1)==max(B)
    p3=B(1);
    if B(2)==min(B)
        p1=B(2);
        p2=B(3);
    else
        p1=B(3);
        p2=B(2);
    end
else
    if B(2)==max(B)
        p3=B(2);
        if B(1)==min(B)
            p1=B(1);
            p2=B(3);
        else
            p1=B(3);
            p2=B(1);
        end
    else 
        p3=B(3);
        if B(1)==min(B)
            p1=B(1);
            p2=B(2);
        else
            p1=B(2);
            p2=B(1);
        end
    end
end
tmax=(p3-p1)/2;
xlswrite(outFile,p1,1,'C77');
xlswrite(outFile,p2,1,'C78');
xlswrite(outFile,p3,1,'C79');
xlswrite(outFile,tmax,1,'C80');

%Margin of safety for A basis
if p1>=0&&p3>0
    MOSR=(sigtstarA/p3)-1;
else
    if p1<0&&p3<=0
        MOSR=(sigcstarA/p1)-1;
    else
        MOSR=min([sigtstarA/p3 sigcstarA/p1])-1;
    end
end
xlswrite(outFile,MOSR,1,'C83');

if p1>=0&&p3>0
    MOST=min([sigtstarA/p3 2*sheartressigA/(p3-p1)])-1;
else
    if p1<0&&p3<=0
        MOST=min([sigcstarA/p1 2*sheartressigA/(p3-p1)])-1;
    else
        MOST=min([sigtstarA/p3 (sigcstarA)/p1 2*sheartressigA/(p3-p1)])-1;
    end
end

xlswrite(outFile,MOST,1,'D83');

if p1>0
    sig1=sigtstarA;
else
    sig1=abs(sigcstarA);
end
if p2>0
    sig2=sigtstarA;
else
    sig2=abs(sigcstarA);
end
if p3>0
    sig3=sigtstarA;
else
    sig3=abs(sigcstarA);
end

SVM=(p1/sig1)^2+(p2/sig2)^2+(p3/sig3)^2-(p1/sig1)*(p2/sig2)-(p1/sig1)*(p3/sig3)-(p2/sig2)*(p3/sig3);
MOSV=1/(SVM^0.5)-1;
xlswrite(outFile,MOSV,1,'E83');

%Margin of safety for B basis
if p1>=0&&p3>0
    MOSRB=(sigtstarB/p3)-1;
else
    if p1<0&&p3<=0
        MOSRB=((sigcstarB)/p1)-1;
    else
        MOSRB=min([sigtstarB/p3 (sigcstarB)/p1])-1;
    end
end
xlswrite(outFile,MOSRB,1,'C84');

if p1>=0&&p3>0
    MOSTB=min([sigtstarB/p3 2*sheartressigB/(p3-p1)])-1;
else
    if p1<0&&p3<=0
        MOSTB=min([sigcstarB/p1 2*sheartressigB/(p3-p1)])-1;
    else
        MOSTB=min([sigtstarB/p3 (sigcstarB)/p1 2*sheartressigB/(p3-p1)])-1;
    end
end

xlswrite(outFile,MOSTB,1,'D84');

if p1>0
    sig1=sigtstarB;
else
    sig1=abs(sigcstarB);
end
if p2>0
    sig2=sigtstarB;
else
    sig2=abs(sigcstarB);
end
if p3>0
    sig3=sigtstarB;
else
    sig3=abs(sigcstarB);
end
SVMB=(p1/sig1)^2+(p2/sig2)^2+(p3/sig3)^2-(p1/sig1)*(p2/sig2)-(p1/sig1)*(p3/sig3)-(p2/sig2)*(p3/sig3);
MOSVB=1/(SVMB^0.5)-1;
xlswrite(outFile,MOSVB,1,'E84');

sigRA=sig*(MOSR+1);
sigRB=sig*(MOSRB+1);
sigTA=sig*(MOST+1);
sigTB=sig*(MOSTB+1);
sigVA=sig*(MOSV+1);
sigVB=sig*(MOSVB+1);
sheRA=she*(MOSR+1);
sheRB=she*(MOSRB+1);
sheTA=she*(MOST+1);
sheTB=she*(MOSTB+1);
sheVA=she*(MOSV+1);
sheVB=she*(MOSVB+1);

xlswrite(outFile,sigRA,1,'C87:C89');
xlswrite(outFile,sheRA,1,'C90:C92');
xlswrite(outFile,sigRB,1,'C95:C97');
xlswrite(outFile,sheRB,1,'C98:C100');
xlswrite(outFile,sigTA,1,'D87:D89');
xlswrite(outFile,sheTA,1,'D90:D92');
xlswrite(outFile,sigTB,1,'D95:D97');
xlswrite(outFile,sheTB,1,'D98:D100');
xlswrite(outFile,sigVA,1,'E87:E89');
xlswrite(outFile,sheVA,1,'E90:E92');
xlswrite(outFile,sigVB,1,'E95:E97');
xlswrite(outFile,sheVB,1,'E98:E100');


