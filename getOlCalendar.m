%% https://stackoverflow.com/questions/40429116/retrieving-outlook-calendar-items-using-matlab
clear all;
close all;

%% Outlook APIに接続
outlook = actxserver('Outlook.Application');
mapi = outlook.GetNamespace('mapi');
explorer = mapi.GetDefaultFolder(9).GetExplorer;
% https://baccholog.com/archives/128
NavModule = explorer.NavigationPane.Modules.GetNavigationModule(1); %予定表を取得
NavGroup = NavModule.NavigationGroups.GetDefaultNavigationGroup(1);% 個人用の予定表を取得
%NavGroup = NavModule.NavigationGroups.GetDefaultNavigationGroup(2);%
%共用の予定表を取得

%% 日付をフォーマット『yyyy/MM/dd HH:mm』または で入力。HH:mmは省略可能
StartDate_MIN = '2020/02/06';
StartDate_MAX = '2020/02/08';


filter = {['[Start] >= ''',strDATstart,''' AND [Start] <= ''', strDATend, '''']};
%filter = {["[Start] >= ''',strDATstart,''' AND [End] <= ''', strDATend, '''"]};

%Startは開始時間、Endは終了時間
%strRestriction : "[Start] <= '02/11/2020 12:00 午前' AND [Start] >= '02/06/2020 12:00 午前'" : String
  
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