%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%                   Patch clamp data analysis
%
%   This script takes data generated from 4 kinds of whole cell patch clamp
%   experiments and pretty plots it according to the data type collected.
%   
%   Data was collected from ClampEx, then exported into an excel spreadsheet,
%   with each experiment for the given patch contained on a different worksheet. The
%   script imports the data from each given worksheet, automatically
%   detects which experiment it came from, and treats it accordingly.
%
%   Experiment types include...
%       1. epsc: spontaneous activity of neuron in I=0, recorded for 60s
%       2. I-clamp: current clamp, response to neuron to injected voltage
%          steps
%       3. IV plot: voltage clamp, response of neuron to injected current.
%          This is then used to calculate voltage gated Na+, K+ fast, and K+
%          slow currents in an I-V plot
%       4. NTX: response of neuron to drug application in I=0. The timing
%          and label information for the drug injections is extracted from a
%          second excel spreadsheet (that may contain multiple worksheets
%          corresponding to different drug treatment experiments) and plotted
%          accordingly.
%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%% specify data path and read designated sheet of excel file
path = '/users/abubnys/Desktop/Matlab_projects/patch_clamp_analysis/';
page = input('page of worksheet? '); % worksheet in excel file
raw_data = xlsread([path 'sample_patch_data.xlsx'],sprintf('Sheet%0.0f',page));

% determining what experiment type applies
epsc = 0;
if size(raw_data) == [8192 9]
    epsc = 1;
end
iclamp = 0;
if size(raw_data) == [8192 8]
    iclamp = 1;
end
iv = 0;
if size(raw_data) == [8192 3]
    iv = 1;
end

% for drug injection, extract information about what drug was injected at
% what time
ntx = 0;
do_label = 0;
if epsc == 0 && iclamp == 0 && iv == 0
    ntx = 1;
    if do_label == 1
        label_page = input('page of label spreadsheet to use? ');
        fil = 'labels';
        [num, txt, raw] = xlsread([path fil],label_page);
        ntx_in = zeros(1,length(txt)-1);
        if length(txt) > 2
            for e = 2:length(txt)
                A = sscanf(txt{e},['%f','%f']);
                ntx_in(e-1) = A(2)*1000;
            end
            if rem(length(txt),2) == 1
                A = textscan(txt{2},'%f \t %f \t %s%s');
                B = strsplit(txt{3});
                r_info = B{4};
            else
                A = textscan(txt{3},'%f \t %f \t %s%s');
                B = strsplit(txt{2},'     ');
                r_infoA = strjoin([A{3} A{4}]);
                r_info = [r_infoA ' ' B{3} ];
            end
        else
            A = strsplit(txt{3},'     ');
            r_info = A{3};
            ntx_in = A{2};
        end
    end
    
    nom = sprintf('ntx_(page%0.0f)',page);
    
    if do_label == 1
        injection_ntx = cell(1,length(txt));
        injection_ntx{1} = '';
        for e = 2:length(txt)
            lin = strsplit(txt{e});
            injection_ntx{e} = lin{4};
        end
    end
end

%% plotting data fromv epsc protocol
% epsc: 8192x9
if epsc == 1
    mv = raw_data(:,1);
    trc = raw_data(:,3:9);
    trc = reshape(trc,1,8192*7);
    
    figure
    hold on
    plot(linspace(0,50,50000),trc(1:50000))
    xlabel('time (s)')
    ylabel('uV')
    title('spontaneous activity')
    %nom = ['epsc(' page ')'];
    nom = sprintf('epsc_(page%0.0f)',page);
    
    rpt = 0;
    while rpt == 0
        yrng = input('y limit range ');
        if isempty(yrng) == 0
            ylim([yrng])
            rpt = input('ok? (y=1, n=0) ');
        else
            rpt = 1;
        end
    end
end

%% plotting data from I-clamp protocol
% Iclamp: 8192x8

if iclamp == 1
    trc = raw_data(:,2:8);
    trc1 = reshape(trc,1,7*8192);
    trc2 = trc1(9501:55000);
    trc3 = reshape(trc2,3500,13);
    
    figure
    plot(trc2)
    set(gca,'XTickLabel',0:5:50)
    xlabel('time (s)')
    ylabel('mV')
    title('I-clamp')
    %nom = ['Iclamp(' page ')'];
    nom = sprintf('Iclamp_(page%0.0f)',page);
    
end

%% plotting data from V-clamp protocol
% Vclamp: 8192x3

if iv == 1
    trc = raw_data(:,2:3);
    trc2 = reshape(trc,1,2*8192);
    trc3 = trc2(3072+73:8816-72);
    trc4 = reshape(trc3,400,14);
    
    figure
    subplot(2,1,1)
    hold on
    plot(linspace(0,0.4,400),trc4(:,2:14))
    xlabel('time (s)')
    ylabel('uV')
    title('V-clamp')
end

%% generating IV plot from V-clamp data
% IV plot
if iv == 1
    Ina_trc = trc4(290:300,2:14);
    Ikt_trc = trc4(290:320,2:14); 
    Iks_trc = trc4(360:380,2:14);
    
    Ikt = max(Ikt_trc);
    Iks = mean(Iks_trc);
    Ina = min(Ina_trc)-Iks;
    Ikf = Ikt-Iks;
    Ikf = Ikf';
    Iks = Iks';
    Ina = Ina';
    %injected = -100:10:150;
    injected = -90:10:30;
    
    subplot(2,1,2)
    hold on
    plot(injected,Ina,injected,Iks,injected,Ikf)
    legend('Na','K slow','K fast','Location','EastOutside')
    xlabel('voltage clamp (mV)')
    ylabel('current (pA)')
    nom = sprintf('IV_(page%0.0f)',page);
end

%% generating plot from drug injection protocol
if ntx == 1
    
    rd = reshape(raw_data,1,size(raw_data,1)*size(raw_data,2));
    plot(rd)
    rpt = 0;
    while rpt == 0
        rng = input('range for data ');
        close
        plot(rng(1):rng(2),rd(rng(1):rng(2)))
        rpt = input('ok? (y=1, n=0) ');
    end
    
    set(gca,'XTick',0:60000:rng(2))
    set(gca,'XTickLabel',0:20)
    xlabel('time (min)')
    ylabel('mV')
    
    rpt = 0;
    while rpt == 0
        yrng = input('y limit range ');
        if isempty(yrng) == 0
            ylim([yrng])
            rpt = input('ok? (y=1, n=0) ');
        else
            rpt = 1;
            yrng = ylim;
        end
    end
    
    %ntx_in = input('NTX injected ');
    hold on
    if do_label == 1
        if length(ntx_in) > 1
            if rem(length(ntx_in),2) == 0
                ntx_legend{1} = '';
                c = 1;
                for e = 1:2:length(ntx_in)
                    plot([ntx_in(e)+10000 ntx_in(e)+10000],yrng,'g','LineWidth',2)
                    ntx_legend{c+1} = injection_ntx{e+1};
                    c = c+1;
                end
                for e = 2:2:length(ntx_in)
                    plot([ntx_in(e)+10000 ntx_in(e)+10000],yrng,'r','LineWidth',2)
                    ntx_legend{c+1} = injection_ntx{e+1};
                    c = c+1;
                end
                legend(ntx_legend)
            else
                plot([ntx_in(1)+10000 ntx_in(1)+10000],yrng,'b','LineWidth',2)
                for e = 2:2:length(ntx_in)
                    plot([ntx_in(e)+10000 ntx_in(e)+10000],yrng,'g','LineWidth',2)
                end
                for e = 3:2:length(ntx_in)
                    plot([ntx_in(e)+10000 ntx_in(e)+10000],yrng,'r','LineWidth',2)
                end
                legend(injection_ntx)
            end
        else
            plot([ntx_in(1)+10000 ntx_in(1)+10000],yrng,'b','LineWidth',2)
            legend(injection_ntx)
        end
    end
    title(r_info)
end

%% saving the figure
tit = input('change title? (leave blank to keep it the same)');
if isempty(tit) == 0
    title(tit)
end
sav = input('save? (1=y,0=n) ');
if sav == 1
    saveas(gcf,[path nom], 'tiff')
    disp('figure saved')
else
    disp('figure not saved')
end

%close all