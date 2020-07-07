function varargout = rank(varargin)
% RANK MATLAB code for rank.fig
%      RANK, by itself, creates a new RANK or raises the existing
%      singleton*.
%
%      H = RANK returns the handle to a new RANK or the handle to
%      the existing singleton*.
%
%      RANK('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in RANK.M with the given input arguments.
%
%      RANK('Property','Value',...) creates a new RANK or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before rank_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to rank_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help rank

% Last Modified by GUIDE v2.5 04-Jul-2020 16:32:59

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @rank_OpeningFcn, ...
                   'gui_OutputFcn',  @rank_OutputFcn, ...
                   'gui_LayoutFcn',  [] , ...
                   'gui_Callback',   []);
if nargin && ischar(varargin{1})
    gui_State.gui_Callback = str2func(varargin{1});
end

if nargout
    [varargout{1:nargout}] = gui_mainfcn(gui_State, varargin{:});
else
    gui_mainfcn(gui_State, varargin{:});
end
% End initialization code - DO NOT EDIT


% --- Executes just before rank is made visible.
function rank_OpeningFcn(hObject, eventdata, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to rank (see VARARGIN)

% Choose default command line output for rank
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes rank wait for user response (see UIRESUME)
% uiwait(handles.figure1);

% --- Outputs from this function are returned to the command line.
function varargout = rank_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;

%--
%function selectionChanged(src,event)
%set(handles.edit9,'Value',src.Value);

% ValueChangedFcn callback
% --- Executes on button press in pushbutton1.
function pushbutton1_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
%global getFileName;
global excelData;
global Head;
global class_name;
global fileName;
global filePath;
set(handles.listbox7,'String','wait...',...
    'Value',1);
[fileName ,filePath] = uigetfile;

if  (isequal(fileName,''))||(isequal(filePath,0))%防止選擇空文件夾
   msgbox('no excel file in the path you selected');
else 
   try
       [excelData,Head] = xlsread(strcat(filePath,'\',fileName));%讀取excel
       [row, col] = find(isnan(excelData)==0);
       for i=1:length(row)
            Head(row(i)+1,col(i)) = {char(string(excelData(row(i),col(i))))};
       end
       class_inx=find(strcmp(Head(1,:),'班級簡稱'));
       class_name = unique(Head(2:end,class_inx));
       set(handles.listbox7,'String',class_name,...
           'Value',1);
       set(handles.text18,'String',strcat(filePath,fileName));
   catch
       msgbox('not correct excel file in the path you selected');
   end
end
%filePath=uigetdir({}); %獲取excel文件存儲目錄
%getFileName=ls(strcat(filePath,'\*.xl*'));  %獲取所選目錄下的文件名
%fileName = cellstr(getFileName); %將string數組轉?cell數組


function edit2_Callback(hObject, eventdata, handles)
% hObject    handle to edit2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit2 as text
%        str2double(get(hObject,'String')) returns contents of edit2 as a double


% --- Executes during object creation, after setting all properties.
function edit2_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit3_Callback(hObject, eventdata, handles)
% hObject    handle to edit3 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit3 as text
%        str2double(get(hObject,'String')) returns contents of edit3 as a double


% --- Executes during object creation, after setting all properties.
function edit3_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit3 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit4_Callback(hObject, eventdata, handles)
% hObject    handle to edit4 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit4 as text
%        str2double(get(hObject,'String')) returns contents of edit4 as a double


% --- Executes during object creation, after setting all properties.
function edit4_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit4 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit5_Callback(hObject, eventdata, handles)
% hObject    handle to edit5 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit5 as text
%        str2double(get(hObject,'String')) returns contents of edit5 as a double


% --- Executes during object creation, after setting all properties.
function edit5_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit5 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit6_Callback(hObject, eventdata, handles)
% hObject    handle to edit6 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit6 as text
%        str2double(get(hObject,'String')) returns contents of edit6 as a double


% --- Executes during object creation, after setting all properties.
function edit6_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit6 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit7_Callback(hObject, eventdata, handles)
% hObject    handle to edit7 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit7 as text
%        str2double(get(hObject,'String')) returns contents of edit7 as a double


% --- Executes during object creation, after setting all properties.
function edit7_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit7 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit8_Callback(hObject, eventdata, handles)
% hObject    handle to edit8 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit8 as text
%        str2double(get(hObject,'String')) returns contents of edit8 as a double


% --- Executes during object creation, after setting all properties.
function edit8_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit8 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in pushbutton2.
function progress_score_avg=progress_score(score_avg,lastscore_avg)
progress_score_avg=[];
for sta=1:size(score_avg,2)
    sta_score_avg = score_avg(:,sta);
    sta_lastscore_avg = lastscore_avg(:,sta);
    %sta_score_avg(find(isnan(sta_score_avg)==1))=[];
    %sta_lastscore_avg(find(isnan(sta_lastscore_avg)==1))=[];
    [~, ~, diffsco, ~, ~]=regress(sta_score_avg,sta_lastscore_avg,0.95);
    %diffsco = score_avg(:,sta)-lastscore_avg(:,sta);
    %sta_diffsco = score_avg(:,sta);
    %sta_diffsco(find(isnan(score_avg(:,sta))~=1))=diffsco;
    progress_score_avg(:,sta) = diffsco;
    %L = size(sta_diff,1);
    %diffsco_avg(find(isnan(diffsco_avg)==1),sta) = 0;
    %r = sqrt( sum(power(diffsco_avg-sum(diffsco_avg)/length(diffsco_avg),2))/L );
    %if r==0
     %   progress_score_avg(:,sta) = ( diffsco_avg-sum(diffsco_avg)/length(diffsco_avg) );
    %else
     %   progress_score_avg(:,sta) = ( diffsco_avg-sum(diffsco_avg)/length(diffsco_avg) )/r;
    %end
end
% hObject    handle to edit8 (see GCBO)


function pushbutton2_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
word = {'金','銀','鳴','銅','響'};
global gold;
global silver;
global copper;
global ring;
global ming;
global limit_average;
global limit_score;
global class_rank;
global group_rank;
global progress_rank;
global excelData;
global Head;
global filePath;
%global chose_class;
global lastHead;
global chose_classlist;
exit_t=0;
gold = str2double(get(handles.edit2,'String'));
silver = str2double(get(handles.edit3,'String'));
copper = str2double(get(handles.edit6,'String'));
ring = str2double(get(handles.edit4,'String'));
ming = str2double(get(handles.edit5,'String'));
limit_average = str2double(get(handles.edit7,'String'));
limit_score = str2double(get(handles.edit8,'String'));
class_rank = get(handles.radiobutton1,'value');
group_rank = get(handles.radiobutton2,'value');
progress_rank = get(handles.radiobutton3,'value');
mode = get(handles.uibuttongroup2,'SelectedObject').String;
group_list = get(handles.popupmenu2,'String');
if  (isempty(Head))&&(isempty(excelData))%防止選擇空文件夾
   msgbox('no excel file in the path you selected');
else 
    if isempty(chose_classlist)
        msgbox('no group  in the your selected');
    else
        try
            if 0 ~= exist(strcat(filePath,'\output'))
                rmdir(strcat(filePath,'\output'), 's');
                mkdir(strcat(filePath,'\output'));
            else
                mkdir(strcat(filePath,'\output'));%新建輸出文件夾
            end
        catch
            exit_t=1;
            set(handles.text17,'String','Do not stay your folder in output path!! or Closed output excel file and run again!!');
            disp 'Do not stay your folder in output path!!'
        end
        a=2;b=2;c=2;d=2;e=2;f=2;
        if exit_t==0
            waiting=waitbar(0,'excuting...,please wait!');%進度條
        end
        for sta_inx=1:length(chose_classlist)
        xlswrite(strcat(filePath,'\output\output.xlsx'),{''}, sta_inx, 'A1');
        end
        es = actxserver('Excel.Application'); % # open Activex server
        ewb = es.Workbooks.Open(strcat(filePath,'\output\output.xlsx')); % # open file (enter full path!)
        for sta_inx=1:length(chose_classlist)
        ewb.Worksheets.Item(sta_inx).Name = strcat(group_list(sta_inx,:),mode,'成績'); % # rename 1st sheet
        end
        ewb.Save % # save to the same file
        ewb.Close(false)
        es.Quit
        for inx_group = 1:length(chose_classlist)
            output_rank={}; 
            output_progress={};
            output_groupprogress={};
            output_classrank={};
            output_grouprank={};
            output_ex={};
            %
            if exit_t==0
                waitbar(double((inx_group-1)/length(chose_classlist)),waiting,strcat('Calculating Group : ',strcat(group_list(1,:),mode,'成績')));
            end
            chose_class = cellstr(chose_classlist{inx_group});
            Data=[];
            lastData=[];
            Data_title=Head{1,:};
            c_index=find(strcmp(Head(1,:),'班級排名'));
            a_index=find(strcmp(Head(1,:),'加權平均'));
            class_inx=find(strcmp(Head(1,:),'班級簡稱'));
            last_index=[find(strcmp(Head(1,:),'班級簡稱')) find(strcmp(Head(1,:),'座號')) find(strcmp(Head(1,:),'姓名'))];
            s_index=[last_index find(strcmp(Head(1,:),'加權平均'))];
            for i=1: length(chose_class) 
                class_data = Head( find(strcmp(Head(:,class_inx),chose_class{i})),:);
                if  i~=1
                    if length(class_data(1,:))~=length(Data(1,:))
                        msgbox('excel file no same group in the path you selected');
                        exit_t=1;
                        break;
                    end
                end
                if progress_rank==1
                    if not(strcmp(get(handles.text16,'String'),'none chose'))
                        lastclass_data = lastHead( find(strcmp(lastHead(:,class_inx),chose_class{i})),:);
                        lastData = [lastData;lastclass_data];
                    else
                        close(ancestor(waiting, 'figure'))
                        msgbox('excel file no last group in the path you selected');
                        exit_t=1;
                        break;
                    end
                end
                if (class_rank==1)&&(exit_t==0)
                    rank=find(strcmp(class_data(:,c_index),'1'));
                    for sta_i=1:length(rank)
                        output_classrank(end+1,:)={class_data{rank(sta_i),s_index},word{5}, ring};
                        output_rank(end+1,:)={class_data{rank(sta_i),s_index},5, ring};
                    end
                    rank=find(strcmp(class_data(:,c_index),'2'));
                    for sta_i=1:length(rank)
                        output_classrank(end+1,:)={class_data{rank(sta_i),s_index},word{5}, ring};
                        output_rank(end+1,:)={class_data{rank(sta_i),s_index},5, ring};
                    end
                    rank=find(strcmp(class_data(:,c_index),'3'));
                    for sta_i=1:length(rank)
                        output_classrank(end+1,:)={class_data{rank(sta_i),s_index},word{5}, ring};
                        output_rank(end+1,:)={class_data{rank(sta_i),s_index},5, ring};
                    end
                end
                Data = [Data;class_data];
            end
            if exit_t==0
                if progress_rank==1 
                    progress_score_avg=[];
                    %score_avg = double(string(Data(:,last_index(end)+1:s_index(end))));
                    %lastscore_avg = double(string(lastData(:,last_index(end)+1:s_index(end))));
                    %progress_score_avg = progress_score(score_avg, lastscore_avg); 
                    progress_score_avg = double(string(lastData(:,end)));
                    output_progress = cellstr([string(lastData(:,1:end-1)) progress_score_avg]);
                    %progressData_title={lastHead{1,[last_index last_index(end)+1:s_index(end)-1 s_index(end)]},'殘排','殘百分比'};
                    progressData_title={lastHead{1,1:end},'殘排','殘百分比'};
                    progress_score_avg(find(isnan(progress_score_avg(:,end))==1),end) = max(progress_score_avg(:,end));
                    [~, ~, forwardRank] = unique(progress_score_avg(:,end)); 
                    reverseRank = max(forwardRank) - forwardRank + 1 ;
                    for sta_i=1:length(reverseRank)
                        back = length(find(reverseRank==sta_i));
                        if back>1
                            reverseRank(find(reverseRank>sta_i))=reverseRank(find(reverseRank>sta_i))+back-1;
                        end    
                    end
                    reverseP = double(reverseRank/length(reverseRank));
                    last=size(output_progress,2)+1;
                    output_progress(:,last:last+1)=num2cell([reverseRank reverseP]);
                    %judge progress price
                    p_index = [1,2,3,5];
                    if class_rank==1
                        for sta_i=1:size(lastData,1)
                            inx=find(strcmp(output_classrank(:,3),lastData(sta_i,3)));
                            if reverseP(sta_i)<=0.03
                                if isempty(inx)
                                    output_rank(end+1,:)={lastData{sta_i,1:end-1},3, ming};
                                else
                                    output_rank(inx,:)={lastData{sta_i,1:end-1},3, ming};
                                    output_classrank(inx,end)={0};
                                end
                                output_groupprogress(end+1,:)={lastData{sta_i,p_index},reverseRank(sta_i),reverseP(sta_i),word{3}, ming};
                            end        
                        end               
                    else
                        for sta_i=1:size(lastData,1)
                            if forwardP(sta_i)<=0.03
                                output_rank(end+1,:)={lastData{sta_i,1:end-1},3, ming};
                                output_groupprogress(end+1,:)={lastData{sta_i,p_index},reverseRank(sta_i),reverseP(sta_i),word{3}, ming};
                            end
                        end     
                    end
                end
                if (class_rank==1) && (group_rank==1)
                    Data_title={Head{1,:},'組排','百分比'};
                    avg_array=double(string(Data(:,a_index)));
                    avg_array(find(isnan(avg_array)==1)) = 0;
                    [~, ~, forwardRank] = unique(avg_array); 
                    reverseRank = max(forwardRank) - forwardRank + 1 ;
                    for sta_i=1:length(reverseRank)
                        back = length(find(reverseRank==sta_i));
                        if back>1
                            reverseRank(find(reverseRank>sta_i))=reverseRank(find(reverseRank>sta_i))+back-1;
                        end    
                    end
                    reverseP = double(reverseRank/length(reverseRank));
                    %[slect_average,slect_inx] = sort(avg_array, 'descend');
                    last=size(Data,2)+1;
                    Data(:,last:last+1)=num2cell([reverseRank reverseP]);
                    for sta_i=1:size(Data,1)
                        pinx=[];
                        if not((any(double(string(Data(sta_i,last_index(end)+1:s_index(end)-1)))<limit_score))||(any(double(string(Data(sta_i,a_index)))<limit_average)))
                            inx=find(strcmp(output_classrank(:,3),Data(sta_i,4)));
                            if progress_rank==1 
                                pinx=find(strcmp(output_groupprogress(:,3),Data(sta_i,4)));
                            end
                            if reverseP(sta_i)<=0.02
                                if not(isempty(pinx))
                                    output_groupprogress(pinx,end)={0};
                                end
                                if isempty(inx)
                                    output_rank(end+1,:)={Data{sta_i,s_index},1, gold};
                                else
                                    output_rank(inx,:)={Data{sta_i,s_index},1, gold};
                                    output_classrank(inx,end)={0};
                                end
                                output_grouprank(end+1,:)={Data{sta_i,s_index},reverseRank(sta_i),reverseP(sta_i),word{1}, gold};
                            elseif reverseP(sta_i)<=0.04
                                if not(isempty(pinx))
                                    output_groupprogress(pinx,end)={0};
                                end
                                if isempty(inx)
                                    output_rank(end+1,:)={Data{sta_i,s_index},2, silver};
                                else
                                    output_rank(inx,:)={Data{sta_i,s_index},2, silver};
                                    output_classrank(inx,end)={0};
                                end    
                                output_grouprank(end+1,:)={Data{sta_i,s_index},reverseRank(sta_i),reverseP(sta_i),word{2}, silver};
                            elseif reverseP(sta_i)<=0.06
                                if isempty(pinx)
                                    if isempty(inx)
                                        output_rank(end+1,:)={Data{sta_i,s_index},4, copper};
                                    else
                                        output_rank(inx,:)={Data{sta_i,s_index},4, copper};
                                        output_classrank(inx,end)={0};
                                    end 
                                    output_grouprank(end+1,:)={Data{sta_i,s_index},reverseRank(sta_i),reverseP(sta_i),word{4}, copper};
                                else
                                    if isempty(inx)
                                        output_grouprank(end+1,:)={Data{sta_i,s_index},reverseRank(sta_i),reverseP(sta_i),word{4}, 0};
                                    else
                                        output_classrank(inx,end)={0};
                                        output_grouprank(end+1,:)={Data{sta_i,s_index},reverseRank(sta_i),reverseP(sta_i),word{4}, 0};
                                    end
                                end
                            end
                        else
                            if reverseP(sta_i)<=0.02
                                output_ex(end+1,:)={Data{sta_i,s_index},reverseRank(sta_i),reverseP(sta_i),word{1}, gold};
                            elseif reverseP(sta_i)<=0.04
                                output_ex(end+1,:)={Data{sta_i,s_index},reverseRank(sta_i),reverseP(sta_i),word{2}, silver};                  
                            elseif reverseP(sta_i)<=0.06
                                output_ex(end+1,:)={Data{sta_i,s_index},reverseRank(sta_i),reverseP(sta_i),word{4}, copper};
                            end              
                        end
                    end      
                elseif group_rank==1
                    Data_title={Head{1,:},'組排','百分比'};
                    avg_array=double(string(Data(:,a_index)));
                    avg_array(find(isnan(avg_array)==1)) = 0;
                    [~, ~, forwardRank] = unique(avg_array); 
                    reverseRank = max(forwardRank) - forwardRank + 1 ;
                    reverseP = double(reverseRank/max(reverseRank));
                    for sta_i=1:length(reverseRank)
                        back = length(find(reverseRank==sta_i));
                        if back>1
                            reverseRank(find(reverseRank>sta_i))=reverseRank(find(reverseRank>sta_i))+back-1;
                        end    
                    end
                    %[slect_average,slect_inx] = sort(avg_array, 'descend');
                    last=size(Data,2)+1;
                    Data(:,last:last+1)=num2cell([reverseRank reverseP]);
                    for sta_i=1:size(Data,1)
                        pinx=[];
                        if not((any(double(string(Data(sta_i,last_index(end)+1:s_index(end)-1)))<limit_score))||(any(double(string(Data(sta_i,a_index)))<limit_average)))
                            if progress_rank==1 
                                pinx=find(strcmp(output_groupprogress(:,3),Data(sta_i,4)));
                            end
                            if reverseP(sta_i)<=0.02
                                if not(isempty(pinx))
                                    output_groupprogress(pinx,end)={0};
                                end
                                output_rank(end+1,:)={Data{sta_i,s_index},1, gold};
                                output_grouprank(end+1,:)={Data{sta_i,s_index},reverseRank(sta_i),reverseP(sta_i),word{1}, gold};
                            elseif reverseP(sta_i)<=0.04
                                if not(isempty(pinx))
                                    output_groupprogress(pinx,end)={0};
                                end
                                output_rank(end+1,:)={Data{sta_i,s_index},2, silver};
                                output_grouprank(end+1,:)={Data{sta_i,s_index},reverseRank(sta_i),reverseP(sta_i),word{2}, silver};               
                            elseif reverseP(sta_i)<=0.06
                                if isempty(pinx)
                                    output_rank(end+1,:)={Data{sta_i,s_index},4, copper};
                                    output_grouprank(end+1,:)={Data{sta_i,s_index},reverseRank(sta_i),reverseP(sta_i),word{4}, copper};
                                else
                                    output_grouprank(end+1,:)={Data{sta_i,s_index},reverseRank(sta_i),reverseP(sta_i),word{4}, 0};
                                end
                            end
                        else
                            if reverseP(sta_i)<=0.02
                                output_ex(end+1,:)={Data{sta_i,s_index},reverseRank(sta_i),reverseP(sta_i),word{1}, gold};
                            elseif reverseP(sta_i)<=0.04
                                output_ex(end+1,:)={Data{sta_i,s_index},reverseRank(sta_i),reverseP(sta_i),word{2}, silver};                  
                            elseif reverseP(sta_i)<=0.06
                                output_ex(end+1,:)={Data{sta_i,s_index},reverseRank(sta_i),reverseP(sta_i),word{4}, copper};
                            end  
                        end
                    end     
                end
                waitbar(double((inx_group-0.5)/length(chose_classlist)),waiting,strcat('Writing Group : ',strcat(group_list(inx_group,:),mode,'成績')));
                try 
                    rank = cell2mat(Data(:,end-1));
                    [~,slect_inx] = sort(rank);
                    Data = Data(slect_inx,:);
                    xlswrite(strcat(filePath,'\output\output.xlsx'),Data_title, strcat(group_list(inx_group,:),mode,'成績'), 'A1');
                    xlswrite(strcat(filePath,'\output\output.xlsx'),Data, strcat(group_list(inx_group,:),mode,'成績'), 'A2');
                    if not(isempty(output_progress))
                        rank = cell2mat(output_progress(:,end-1));
                        [~,slect_inx] = sort(rank);
                        output_progress = output_progress(slect_inx,:);
                        xlswrite(strcat(filePath,'\output\output.xlsx'),progressData_title, '進步', 'A1');
                        xlswrite(strcat(filePath,'\output\output.xlsx'),output_progress,'進步', strcat('A',num2str(a)));%將讀取的工作表數字寫入excel
                        a = a + size(output_progress,1);
                    end
                    if not(isempty(output_groupprogress))
                        rank = cell2mat(output_groupprogress(:,end-3));
                        [~,slect_inx] = sort(rank);
                        output_groupprogress = output_groupprogress(slect_inx,:);
                        xlswrite(strcat(filePath,'\output\output.xlsx'),{Head{1,s_index(1:3)},'殘差平均','殘排','殘百分比','獎項','金額'}, '鳴', 'A1');%將讀取的工作表表頭寫入excel
                        xlswrite(strcat(filePath,'\output\output.xlsx'),output_groupprogress,'鳴', strcat('A',num2str(b)));%將讀取的工作表數字寫入excel
                        b = b + size(output_groupprogress,1);
                    end 
                    if not(isempty(output_grouprank))
                        rank = cell2mat(output_grouprank(:,end-3));
                        [~,slect_inx] = sort(rank);
                        output_grouprank = output_grouprank(slect_inx,:);
                        xlswrite(strcat(filePath,'\output\output.xlsx'),{Head{1,s_index},'組排','百分比','獎項','金額'}, '金銀銅', 'A1');%將讀取的工作表表頭寫入excel                            
                        xlswrite(strcat(filePath,'\output\output.xlsx'),output_grouprank,'金銀銅', strcat('A',num2str(c)));%將讀取的工作表數字寫入excel
                        c = c + size(output_grouprank,1);
                    end
                    if not(isempty(output_classrank))
                        xlswrite(strcat(filePath,'\output\output.xlsx'),{Head{1,s_index},'獎項','金額'}, '響', 'A1');%將讀取的工作表表頭寫入excel                          
                        xlswrite(strcat(filePath,'\output\output.xlsx'),output_classrank,'響', strcat('A',num2str(d)));%將讀取的工作表數字寫入excel
                        d = d + size(output_classrank,1);
                    end  
                    if not(isempty(output_rank))
                        gn = length(find(cell2mat(output_rank(:,end-1))==1));
                        sn = length(find(cell2mat(output_rank(:,end-1))==2));
                        mn = length(find(cell2mat(output_rank(:,end-1))==3));
                        cn = length(find(cell2mat(output_rank(:,end-1))==4));
                        rn = length(find(cell2mat(output_rank(:,end-1))==5));
                        xlswrite(strcat(filePath,'\output\output.xlsx'),{'獎項', '金額', '人數', '總額'}, '清冊', 'H3');
                        xlswrite(strcat(filePath,'\output\output.xlsx'),{'金', gold, gn, gn*gold}, '清冊', 'H4');
                        xlswrite(strcat(filePath,'\output\output.xlsx'),{'金', silver, sn, sn*silver}, '清冊', 'H5');
                        xlswrite(strcat(filePath,'\output\output.xlsx'),{'金', ming, mn, mn*ming}, '清冊', 'H6');
                        xlswrite(strcat(filePath,'\output\output.xlsx'),{'金', copper, cn, cn*copper}, '清冊', 'H7');
                        xlswrite(strcat(filePath,'\output\output.xlsx'),{'金', ring, rn, rn*ring}, '清冊', 'H8');
                        xlswrite(strcat(filePath,'\output\output.xlsx'),{' ', ' ', gn+sn+mn+cn+rn, gn*gold+sn*silver+mn*ming+cn*copper+rn*ring}, '清冊', 'H9');
                        xlswrite(strcat(filePath,'\output\output.xlsx'),{' ', ' ', length(find(cell2mat(output_rank(:,end-1))==[1,2,3,4,5])), length(find(cell2mat(output_rank(:,end-1))==5))*ring}, '清冊', 'H10');
                        output_rank(:,end-1) = word(cell2mat(output_rank(:,end-1)))';
                        output_rank(:,4)=[];
                        xlswrite(strcat(filePath,'\output\output.xlsx'),{Head{1,last_index},'獎項','金額'}, '清冊', 'A1');%將讀取的工作表表頭寫入excel
                        xlswrite(strcat(filePath,'\output\output.xlsx'),output_rank,'清冊', strcat('A',num2str(e)));%將讀取的工作表數字寫入excel
                        e = e + size(output_rank,1);
                    end
                    if not(isempty(output_ex))
                        xlswrite(strcat(filePath,'\output\output.xlsx'),{Head{1,s_index},'組排','百分比','獎項','金額'}, '不符合資格', 'A1');%將讀取的工作表表頭寫入excel
                        xlswrite(strcat(filePath,'\output\output.xlsx'),output_ex,'不符合資格', strcat('A',num2str(f)));%將讀取的工作表數字寫入excel
                        f = f + size(output_ex,1);
                    end
                    waitbar(double(inx_group/length(chose_classlist)),waiting,strcat('Writing Group : ',strcat(group_list(inx_group,:),mode,'成績')));
                catch
                    exit_t=1;
                    close(waiting)%關閉進度條
                    set(handles.text17,'String','Please closed output excel file and run again!!');
                    disp 'Please closed output excel file and run again!!'
                end
            end 
        end
        if exit_t==0
            waitbar(1,waiting,strcat('Loaded Successful'));
            disp 'you can find output file there:'
            outputPath = strcat(filePath,'\output\output.xlsx');
            set(handles.text17,'String',strcat('finished,get output file in:',outputPath));
            msgbox(strcat('finished,get output file in:',outputPath),'Success','Help');%prompt message
            waitbar(1,waiting,strcat('Finish all group ',mode,'成績'));
            close(waiting)%關閉進度條
        end
    end
end

% --- Executes on button press in radiobutton1.
function radiobutton1_Callback(hObject, eventdata, handles)
% hObject    handle to radiobutton1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hint: get(hObject,'Value') returns toggle state of radiobutton1


% --- Executes on button press in radiobutton2.
function radiobutton2_Callback(hObject, eventdata, handles)
% hObject    handle to radiobutton2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hint: get(hObject,'Value') returns toggle state of radiobutton2


% --- Executes on button press in radiobutton3.
function radiobutton3_Callback(hObject, eventdata, handles)
% hObject    handle to radiobutton3 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hint: get(hObject,'Value') returns toggle state of radiobutton3


% --- Executes on selection change in listbox4.
function listbox4_Callback(hObject, eventdata, handles)
% hObject    handle to listbox4 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns listbox4 contents as cell array
%        contents{get(hObject,'Value')} returns selected item from listbox4


% --- Executes during object creation, after setting all properties.
function listbox4_CreateFcn(hObject, eventdata, handles)
% hObject    handle to listbox4 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: listbox controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit9_Callback(hObject, eventdata, handles)
% hObject    handle to edit9 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit9 as text
%        str2double(get(hObject,'String')) returns contents of edit9 as a double


% --- Executes during object creation, after setting all properties.
function edit9_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit9 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in listbox7.
function listbox7_Callback(hObject, eventdata, handles)
% hObject    handle to listbox7 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
global chose_class
contents = get(hObject,'String');
inx = get(hObject,'Value');
chose_class = get(handles.listbox8,'String');
if strcmp(chose_class,'None')
    chose_class = contents(inx);
elseif not(any(strcmp(chose_class,contents(inx))))
    chose_class = [chose_class;contents(inx)];
end
set(handles.listbox8,'String',chose_class,...
    'Value',1);
% Hints: contents = cellstr(get(hObject,'String')) returns listbox7 contents as cell array
%        contents{get(hObject,'Value')} returns selected item from listbox7


% --- Executes during object creation, after setting all properties.
function listbox7_CreateFcn(hObject, eventdata, handles)
% hObject    handle to listbox7 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: listbox controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in pushbutton3.
function pushbutton3_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton3 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)


% --- Executes on selection change in listbox8.
function listbox8_Callback(hObject, eventdata, handles)
% hObject    handle to listbox8 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
global chose_class
contents = get(hObject,'String');
inx = get(hObject,'Value');
if not(strcmp(contents,'None'))
    chose_class = chose_class(find(not(strcmp(chose_class,contents(inx)))));
else
    chose_class={'None'};
end
set(handles.listbox8,'String',chose_class,...
    'Value',1);
% Hints: contents = cellstr(get(hObject,'String')) returns listbox8 contents as cell array
%        contents{get(hObject,'Value')} returns selected item from listbox8


% --- Executes during object creation, after setting all properties.
function listbox8_CreateFcn(hObject, eventdata, handles)
% hObject    handle to listbox8 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: listbox controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in pushbutton4.
function pushbutton4_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton4 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
global lastexcelData;
global lastHead;
[fileName ,filePath] = uigetfile;
if  isequal(fileName,'')||(isequal(filePath,0))%防止選擇空文件夾
   msgbox('no excel file in the path you selected');
else 
   try
       [lastexcelData,lastHead] = xlsread(strcat(filePath,'\',fileName));%讀取excel
       [row, col] = find(isnan(lastexcelData)==0);
       for i=1:length(row)
            lastHead(row(i)+1,col(i)) = {char(string(lastexcelData(row(i),col(i))))};
       end
       set(handles.text16,'String',strcat(filePath,fileName));
   catch
       msgbox('not correct excel file in the path you selected');
   end
end


% --- Executes on selection change in popupmenu2.
function popupmenu2_Callback(hObject, eventdata, handles)
% hObject    handle to popupmenu2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
global chose_classlist;
set(handles.listbox9,'String',chose_classlist{1,get(handles.popupmenu2,'Value')},...
'Value',1);
% Hints: contents = cellstr(get(hObject,'String')) returns popupmenu2 contents as cell array
%        contents{get(hObject,'Value')} returns selected item from popupmenu2


% --- Executes during object creation, after setting all properties.
function popupmenu2_CreateFcn(hObject, eventdata, handles)
% hObject    handle to popupmenu2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in listbox9.
function listbox9_Callback(hObject, eventdata, handles)
% hObject    handle to listbox9 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns listbox9 contents as cell array
%        contents{get(hObject,'Value')} returns selected item from listbox9


% --- Executes during object creation, after setting all properties.
function listbox9_CreateFcn(hObject, eventdata, handles)
% hObject    handle to listbox9 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: listbox controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in pushbutton7.
function pushbutton7_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton7 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
global chose_classlist;
chose_class=get(handles.listbox8,'String');
set(handles.listbox8,'String','None');
group=get(handles.popupmenu2,'String');
if (isempty(chose_class))||(any(strcmp(chose_class,'None')))
    if strcmp(group,'none group')
        chose_classlist={};
    end
    msgbox('no class in your selected');
else
    new_group = get(handles.edit11,'String');
    if strcmp(group,'none group')
        if length(chose_classlist)~=length(group)
            chose_classlist={};
            set(handles.popupmenu2,'String','none group');
            set(handles.listbox9,'String','None',...
            'Value',1);
        end
        group=new_group;
        gvalue=size(group,1);
    else
        inx = find(strcmp(string(group),new_group));
        if isempty(inx)
            group = char(group,new_group);
            gvalue=size(group,1);
        end
    end
    if isempty(chose_classlist)
        chose_classlist={string(chose_class)};
    else
        if not(isempty(inx))
            gvalue=inx;
            chose_classlist(1,inx)={string(chose_class)};
        else
            chose_classlist={chose_classlist{:,:},string(chose_class)};
        end
    end
    set(handles.popupmenu2,'Value',gvalue);
    set(handles.popupmenu2,'String',group);
    set(handles.listbox9,'String',chose_class,...
    'Value',1);
end

% --- Executes on button press in pushbutton9.
function pushbutton9_Callback(~, eventdata, handles)
% hObject    handle to pushbutton9 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
global chose_classlist;
group_list = get(handles.popupmenu2,'String');
class_list = get(handles.listbox9,'String');
group = get(handles.popupmenu2,'Value');
class = get(handles.listbox9,'Value');
if length(chose_classlist)>1
    if length(chose_classlist{1,group})>1
        state_class = chose_classlist{1,group};
        state_class(class:end-1,1) = chose_classlist{1,group}(class+1:end,1);
        chose_classlist(1,group) = {string(state_class(1:end-1,1))};
        set(handles.listbox9,'String',chose_classlist{1,group},...
        'Value',1);
    else
        chose_classlist(group) = [];
        group_list(group,:)=[];
        set(handles.popupmenu2,'String',group_list);
        set(handles.popupmenu2,'Value',1);
        set(handles.listbox9,'String',chose_classlist{1,1},...
        'Value',1);
    end
elseif length(chose_classlist)==1
    if length(chose_classlist{1,group})>1
        state_class = chose_classlist{1,group};
        state_class(class:end-1,1) = chose_classlist{1,group}(class+1:end,1);
        chose_classlist(1,group) = {string(state_class(1:end-1,1))};
        set(handles.listbox9,'String',chose_classlist{1,group},...
        'Value',1);
    else
        chose_classlist(group) = [];
        group_list(group,:)=[];
        set(handles.popupmenu2,'String','none group');
        set(handles.popupmenu2,'Value',1);   
        set(handles.listbox9,'String','None',...
        'Value',1);
    end
end




function edit11_Callback(hObject, eventdata, handles)
% hObject    handle to edit11 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit11 as text
%        str2double(get(hObject,'String')) returns contents of edit11 as a double


% --- Executes during object creation, after setting all properties.
function edit11_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit11 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
