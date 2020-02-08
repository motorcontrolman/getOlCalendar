%% �͂��߂�
% �l��s���̍����A�w�Ɩ����ׂ̌����鉻�x�����߂��Ă͂��܂��񂩁B
% 
% ���̂ɐE�ꂪ�Õ��ŁA�v���W�F�N�g�Ǘ��c�[���ȂǓ�������Ă͂��Ȃ��B
% 
% ���邢�͓������ꂽ�v���W�F�N�g�Ǘ��c�[�����C�P�ĂȂ��āA���񂺂�蒅���Ă��Ȃ��B
% 
% ����Ȏ��A���Ȃ��Ȃ�ǂ����܂����H
% 
% 
% 
% ���u���[��A�\���Outlook�ŊǗ����Ă邩��A���̃f�[�^�𕉉ׂ̌����鉻�ɂ��g���Ȃ����̂��c�v
% 
% ���u�c����IMATLAB�ŉ��������炦�����I�v
% 
% ���u�e�L�X�g�f�B�[�v���[�j���O�g���Ή����o������A<https://qiita.com/aoimidori/items/796db2e0ce90f64f30d1 
% Advent Calendar>�ɂ��������̂��������B�v
% 
% 
% 
% �Ƃ������ƂŁAMATLAB�ŉ������Ă����܂��B
% 
% �܂��{�e�ł́AOutlook����\��\���Ԃ�������MATLAB�Ɏ����Ă���܂ł������܂��B
% 
% 
% 
% �Ȃ��A�R�[�h�̑唼�͉��LURL���Q�l�ɍ쐬���Ă��܂��B
% 
% ���̂܂܂��Ɠ����Ȃ������̂Ń����e��������M�����肵�܂����B
% 
% <https://stackoverflow.com/questions/40429116/retrieving-outlook-calendar-items-using-matlab 
% https://stackoverflow.com/questions/40429116/retrieving-outlook-calendar-items-using-matlab>
% 
% 
% 
% �܂����̋L���͈�コ���livescript2markdown���g���č쐬���Ă��܂��A�֗��I
% 
% <https://github.com/minoue-xx/livescript2markdown https://github.com/minoue-xx/livescript2markdown>
% 
% 

clear all;
close all;
%% 1. Outlook API�ɐڑ�
% ���L�R�}���h�����s���邱�ƂŁAMATLAB����Outlook API�ɐڑ����܂��B

outlook = actxserver('Outlook.Application');
mapi = outlook.GetNamespace('mapi');
%% 
% ���ɁAGetDefaultFolderk���\�b�h��@����Outlook�̏���̃t�H���_�ɃA�N�Z�X���܂��B
% 
% ���\�b�h�̈����ƃt�H���_�̑Ή��\�͉��LURL���Q�l�A�\��\�t�H���_�́u9�v�ɂ���܂��B
% 
% <https://baccholog.com/archives/128 Baccho Log [WSH]Outlook�̑���>

explorer = mapi.GetDefaultFolder(9).GetExplorer;
%% 
% ����Ɍ@�艺���Ă������ƂŁA�l�p�⋤�p�̗\��\���擾���邱�Ƃ��o���܂��B
% 
% ����͌l�p�̗\��\���擾���Ă݂܂��傤�B

NavModule = explorer.NavigationPane.Modules.GetNavigationModule(1); %�\��\���擾
NavGroup = NavModule.NavigationGroups.GetDefaultNavigationGroup(1);% �l�p�̗\��\���擾
%NavGroup = NavModule.NavigationGroups.GetDefaultNavigationGroup(2);% ���p�̗\��\���擾
%% 2.����̓����Ńt�B���^����
% �wyyyy/MM/dd HH:mm�x�܂��́wMM/dd/yyyy HH:mm�x�̃t�H�[�}�b�g�œ��t���w�肵�܂��B
% 
% HH:mm�͏ȗ��\�ł����i���̏ꍇAM8:00���炢�Ɏw��H�j�ςȌ��ʂ��o�Ȃ��悤�ȗ����Ȃ��̂��g�ł��B
% 
% �܂��t�B���^�͗\��̊J�n���Ԃ��邢�͏I�����ԂŐݒ�ł��܂��B�J�n���ԂŃt�B���^����ꍇ�͉��L�B

StartDate_MIN = '2020/02/05 00:00';
StartDate_MAX = '2020/02/06 23:00';

%[Start]�͊J�n���ԁA[End]�͏I�����Ԃ��t�B���^�����Ƃ��Ďg�p
filter = {['[Start] >= ''',StartDate_MIN,''' AND [Start] <= ''', StartDate_MAX, '''']};
%% 3.�\����擾�AMATLAB�̃e�[�u���ɐ��^����
% �擾�̑ΏۂƂȂ�Outlook�\��\�͉��L�ł��B
% 
% �t�B���^������2/5~2/6�Ɏw�肵���̂ŁA�u�Ƒ��ƒ����͂�v�ȍ~�̗\�肪����ΐ����ł��B
% 
% 
% 
% 
% 
% �擾����уe�[�u���̐��^�͉��L�B
% 
% �e�[�u�����^�͐F�X�Ȃ���������Ƃ͎v���܂����A����̏ꍇ�͂��ꂪ��ԉǐ����������ȂƁB

for i=1:NavGroup.NavigationFolders.Count
    NavFolder = NavGroup.NavigationFolders.Item(i);
    LST = NavFolder.Folder.Items;
    %LST.IncludeRecurrences = -1;
    LST.Sort('[Start]');
    LST_Restrict = LST.Restrict(filter{1});
    Cnt = LST_Restrict.Count;
    sz = [Cnt 4];
    varNames = {'Subject','Location','Start','End'};
    varTypes = {'string','string','datetime','datetime'};
    Calendar_Table = table('Size',sz,'VariableTypes',varTypes,'VariableNames',varNames);
    for j = 1:Cnt
        Calendar_Table.Subject(j) = LST_Restrict.Item(j).Subject;
        Calendar_Table.Location(j) = LST_Restrict.Item(j).Location;
        Calendar_Table.Start(j) = LST_Restrict.Item(j).Start;
        Calendar_Table.End(j) = LST_Restrict.Item(j).End;
    end
end

Calendar_Table
%% 
% ��������B
%% 4.�\��̒������v�Z����
% �ŏI�S�[�����u�Ɩ����ׂ̌����鉻�v�Ȃ̂ŁA���ꂼ��̗\��̒������m�肽���ł��ˁB
% 
% �\��̒����́Afor���[�v���ɂ����� `LST_Restrict.Item(j).Duration` �Ŏ��o���܂����AMATLAB�̃e�[�u���^�֗̕������������߂ɂ����ăe�[�u�����Ōv�Z�����܂��B�R�[�h�͉��L�B

Calendar_Table.Duration = Calendar_Table.End - Calendar_Table.Start
%% 
% �֗����Ȃ�(����)
%% 5.���ׂ������鉻����
% �ŏI�S�[����Subject����e�L�X�g�f�B�[�v���[�j���O�ŋƖ����e��ސ��A�Ɩ����ׂ������鉻���邱�Ƃł��B���A�����Location��p���Ăǂ��ɂǂꂾ�������̂��������鉻����ɗ��߂܂��B�e�L�X�g�f�B�[�v���[�j���O�F�X��ς����Ȃ̂ŁB
% 
% Location�ɂ͂��炩���ߎ���A��ЁA�ړ��̂����ꂩ���܂ނ悤�ɂ��Ă������̂ŁA<https://jp.mathworks.com/help/matlab/ref/contains.html 
% contains�֐�>���g���ăf�[�^�������čs���܂��B

Location_Name = ["����";"���";"�ړ�"];
Result_Table = table(Location_Name);
Duration_sum = [];
for k = 1:length(Result_Table.Location_Name)
    tmp = contains(Calendar_Table.Location , Result_Table.Location_Name(k)) .* Calendar_Table.Duration;
    Duration_sum = [Duration_sum; sum(tmp)];
end

Result_Table = table(Location_Name, Duration_sum)
pie(hours(Result_Table.Duration_sum), Result_Table.Location_Name);
%% 
% ����́c����܂蓭���Ȃ��l�ł��ˁB
%% ������
% �Ɩ��������c�[���A���ɏo����Ă�����̂ŏ\���ړI�͉ʂ�����Ǝv���܂��B���A�����ō��΃s���|�C���g�Ō��ʂ̂�����̂��ł��Ă�낵�����ƁB�����ĉ����y�����ł��B