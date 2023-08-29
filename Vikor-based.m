clc
clear all                          %Å«ò ò—œ‰ Å‰Ã—Â ›—„«‰ Ê Õ«›ŸÂ%
close all
format shortG
 
%%
V=0.5;                              % ⁄ÌÌ‰ ‘«Œ’ v («ê—  Ê«›ﬁ »« «òÀ—Ì  ¬—«¡ »«‘œ=0.5)%

A=xlsread('mcdm.xlsx',1);           %›—«Œ«‰Ì «ÿ·«⁄«  «ò”·%

W=A(1,:);                           % ⁄ÌÌ‰ —œÌ› Ê“‰ Â«%

DM=A(3:end,:);                      %¬œ—” „« —Ì”  ’„Ì„%

[na,nc]=size(DM);                   %«‰œ«“Â „« —Ì”  ’„Ì„%





%%  
FP=zeros(1,nc);                  % „« —Ì” »—«Ì ‰êÂ œ«—Ì „ﬁ«œÌ— «ÌœÂ ¬· „À»  Ê „‰›Ì(—«»ÿÂ 4)%            
FN=zeros(1,nc);

for c=1:nc
    
    if W(c)>0
    FP(c)=max(DM(:,c));                       %«ê— «»⁄ «“ ‰Ê⁄ „À» (”Êœ) »«‘œ(—«»ÿÂ4)% 
    FN(c)=min(DM(:,c));    
    else
    FP(c)=min(DM(:,c));                        %«ê—  «»⁄ «“ ‰Ê⁄ „‰›Ì (Â“Ì‰Â) »«‘œ (—«»ÿÂ4)%
    FN(c)=max(DM(:,c));
    end
    

end



%%  

W=abs(W);         % Ê“‰ Ê œ‰»«·Â «‘%       

S=zeros(1,na);        %„« —Ì” »—«Ì ‰êÂ œ«—Ì „ﬁ«œÌ— RÊS(—«»ÿÂ5)%
R=zeros(1,na);

for a=1:na
    
E=zeros(1,nc);

for c=1:nc
    
    p=W(c)*(FP(c)-DM(a,c))/(FP(c)-FN(c));     % ⁄Ì‰ „ﬁœ«— ”Êœ„‰œÌ (S) Ê  «”›(R) ÿ»ﬁ —«»ÿÂ 5%
    E(c)=p;       
    
end

S(a)=sum(E);
R(a)=max(E);

end



%%  

SP=min(S);      %„ﬁ«œÌ— *s Ê -s ÿ»ﬁ —«»ÿÂ 6%
SN=max(S);

RP=min(R);      %„ﬁœÌ— *RÊ -R ÿ»ﬁ —«»ÿÂ 6%
RN=max(R);

Q=zeros(1,na);    %„« —Ì” »—«Ì ‰êÂ œ«—Ì „ﬁ«œÌ—Q%


for a=1:na
    
    
s=(SP-S(a))/(SP-SN);        %„Õ«”»Â „ﬁœ«—Q »—«Ì Â— ê“Ì‰Â(—«»ÿÂ6)%
r=(RP-R(a))/(RP-RN);

Q(a)=V*s+(1-V)*r;

end





%%                                              
                                                                                       %‰„«Ì‘ ‰ «ÌÃ%

disp('****************** Q ******************')
[value,index]=sort(Q);                                                                     %„— » ò—œ‰ „ﬁ«œÌ— Q%

for a=1:na 
disp(['Rank' num2str(a) ' Alter' num2str(index(a)) ' Q = ' num2str(value(a))])              %‰„«Ì‘ „ﬁ«œÌ— Q%
end
disp(' ')


disp('****************** S ******************')            
                                                                                          %„— » ò—œ‰ „ﬁ«œÌ— S%
[value,index]=sort(S);

for a=1:na 
disp(['Rank' num2str(a) ' Alter' num2str(index(a)) ' S = ' num2str(value(a))])            %‰„«Ì‘ „ﬁ«œÌ— S%
end
disp(' ')


disp('****************** R ******************')
[value,index]=sort(R);                                                                    %„— » ò—œ‰ „ﬁ«œÌ— R%

for a=1:na 
disp(['Rank' num2str(a) ' Alter' num2str(index(a)) ' R = ' num2str(value(a))])             %‰„«Ì‘ „ﬁ«œÌ—R%
end
disp(' ')


