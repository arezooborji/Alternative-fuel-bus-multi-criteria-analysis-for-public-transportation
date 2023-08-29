clc
clear all                          %�ǘ ���� ����� ����� � �����%
close all
format shortG
 
%%
V=0.5;                              %����� ���� v (ǐ� ����� �� ǘ���� ���� ����=0.5)%

A=xlsread('mcdm.xlsx',1);           %������� ������� ǘ��%

W=A(1,:);                           %����� ���� ��� ��%

DM=A(3:end,:);                      %���� ������ �����%

[na,nc]=size(DM);                   %������ ������ �����%





%%  
FP=zeros(1,nc);                  % ������ ���� �� ���� ������ ���� �� ���� � ����(����� 4)%            
FN=zeros(1,nc);

for c=1:nc
    
    if W(c)>0
    FP(c)=max(DM(:,c));                       %ǐ����� �� ��� ����(���) ����(�����4)% 
    FN(c)=min(DM(:,c));    
    else
    FP(c)=min(DM(:,c));                        %ǐ� ���� �� ��� ���� (�����) ���� (�����4)%
    FN(c)=max(DM(:,c));
    end
    

end



%%  

W=abs(W);         % ��� � ������ ��%       

S=zeros(1,na);        %������ ���� �� ���� ������ R�S(�����5)%
R=zeros(1,na);

for a=1:na
    
E=zeros(1,nc);

for c=1:nc
    
    p=W(c)*(FP(c)-DM(a,c))/(FP(c)-FN(c));     %���� ����� ������� (S) � ����(R) ��� ����� 5%
    E(c)=p;       
    
end

S(a)=sum(E);
R(a)=max(E);

end



%%  

SP=min(S);      %������ *s � -s ��� ����� 6%
SN=max(S);

RP=min(R);      %����� *R� -R ��� ����� 6%
RN=max(R);

Q=zeros(1,na);    %������ ���� �� ���� ������Q%


for a=1:na
    
    
s=(SP-S(a))/(SP-SN);        %������ �����Q ���� �� �����(�����6)%
r=(RP-R(a))/(RP-RN);

Q(a)=V*s+(1-V)*r;

end





%%                                              
                                                                                       %����� �����%

disp('****************** Q ******************')
[value,index]=sort(Q);                                                                     %���� ���� ������ Q%

for a=1:na 
disp(['Rank' num2str(a) ' Alter' num2str(index(a)) ' Q = ' num2str(value(a))])              %����� ������ Q%
end
disp(' ')


disp('****************** S ******************')            
                                                                                          %���� ���� ������ S%
[value,index]=sort(S);

for a=1:na 
disp(['Rank' num2str(a) ' Alter' num2str(index(a)) ' S = ' num2str(value(a))])            %����� ������ S%
end
disp(' ')


disp('****************** R ******************')
[value,index]=sort(R);                                                                    %���� ���� ������ R%

for a=1:na 
disp(['Rank' num2str(a) ' Alter' num2str(index(a)) ' R = ' num2str(value(a))])             %����� ������R%
end
disp(' ')


