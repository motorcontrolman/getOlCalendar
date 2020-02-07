%% https://stackoverflow.com/questions/40429116/retrieving-outlook-calendar-items-using-matlab
clear all;
close all;

%% Outlook API�ɐڑ�
outlook = actxserver('Outlook.Application');
mapi = outlook.GetNamespace('mapi');
explorer = mapi.GetDefaultFolder(9).GetExplorer;
% https://baccholog.com/archives/128
NavModule = explorer.NavigationPane.Modules.GetNavigationModule(1); %�\��\���擾
NavGroup = NavModule.NavigationGroups.GetDefaultNavigationGroup(1);% �l�p�̗\��\���擾
%NavGroup = NavModule.NavigationGroups.GetDefaultNavigationGroup(2);%
%���p�̗\��\���擾

%% ���t���t�H�[�}�b�g�wyyyy/MM/dd HH:mm�x�܂��� �œ��́BHH:mm�͏ȗ��\
StartDate_MIN = '2020/02/06';
StartDate_MAX = '2020/02/08';


filter = {['[Start] >= ''',strDATstart,''' AND [Start] <= ''', strDATend, '''']};
%filter = {["[Start] >= ''',strDATstart,''' AND [End] <= ''', strDATend, '''"]};

%Start�͊J�n���ԁAEnd�͏I������
%strRestriction : "[Start] <= '02/11/2020 12:00 �ߑO' AND [Start] >= '02/06/2020 12:00 �ߑO'" : String
  
for i=1:NavGroup.NavigationFolders.Count
    NavFolder = NavGroup.NavigationFolders.Item(i);
    LST = NavFolder.Folder.Items;
    %LST.IncludeRecurrences = -1;
    LST.Sort('[Start]');
    LST_Restrict = LST.Restrict(filter{1});
    Cnt = LST_Restrict.Count;
    sz = [Cnt 3];
    varNames = {'Subject','Start','End'};
    varTypes = {'string','datetime','datetime'};
    Calendar_Table = table('Size',sz,'VariableTypes',varTypes,'VariableNames',varNames);
    for j = 1:Cnt
        Calendar_Table.Subject(j) = LST_Restrict.Item(j).Subject;
        Calendar_Table.Start(j) = LST_Restrict.Item(j).Start;
        Calendar_Table.End(j) = LST_Restrict.Item(j).End;
    end
end

Calendar_Table.Dulation = Calendar_Table.End - Calendar_Table.Start