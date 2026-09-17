using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;

[assembly: AssemblyTitle("Lab Photo Tools 설치")]
[assembly: AssemblyVersion("0.1.17.0")]
[assembly: AssemblyFileVersion("0.1.17.0")]
internal static class SetupCore
{
    internal const string Version="0.1.17";
    internal static string PrerequisiteError()
    {
        if(!Environment.Is64BitOperatingSystem || String.Equals(Environment.GetEnvironmentVariable("PROCESSOR_ARCHITECTURE"),"ARM64",StringComparison.OrdinalIgnoreCase) || String.Equals(Environment.GetEnvironmentVariable("PROCESSOR_ARCHITEW6432"),"ARM64",StringComparison.OrdinalIgnoreCase))
            return "Intel/AMD 64비트 Windows 10 또는 11이 필요합니다.";
        if(Environment.OSVersion.Version.Major<10)return "Windows 10 또는 11이 필요합니다.";
        bool framework=false,powerpoint=false;
        foreach(RegistryView view in new[]{RegistryView.Registry32,RegistryView.Registry64})
        {
            using(RegistryKey machine=RegistryKey.OpenBaseKey(RegistryHive.LocalMachine,view))
            using(RegistryKey key=machine.OpenSubKey(@"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full"))
                if(key!=null&&Convert.ToInt32(key.GetValue("Release",0))>=528040)framework=true;
            using(RegistryKey classes=RegistryKey.OpenBaseKey(RegistryHive.ClassesRoot,view))
            using(RegistryKey key=classes.OpenSubKey(@"PowerPoint.Application\CLSID"))
                if(key!=null&&!String.IsNullOrWhiteSpace(Convert.ToString(key.GetValue(""))))powerpoint=true;
        }
        if(!framework)return ".NET Framework 4.8 이상이 필요합니다. Windows 업데이트 후 다시 실행하세요.";
        if(!powerpoint)return "Windows 데스크톱 PowerPoint를 먼저 설치하세요. 웹 PowerPoint에서는 이 추가 기능을 사용할 수 없습니다.";
        if(Process.GetProcessesByName("POWERPNT").Length>0)return "프레젠테이션을 저장하고 모든 PowerPoint 창을 닫은 다음 다시 설치하세요.";
        return null;
    }
    internal static string ExtractPayload(string directory)
    {
        if(Directory.Exists(directory))throw new IOException("설치 임시 폴더가 이미 있습니다.");
        Directory.CreateDirectory(directory);
        string prefix=Path.GetFullPath(directory).TrimEnd(Path.DirectorySeparatorChar)+Path.DirectorySeparatorChar;
        using(Stream stream=Assembly.GetExecutingAssembly().GetManifestResourceStream("LabPhotoTools.payload.zip"))
        {
            if(stream==null)throw new IOException("설치 파일에 구성 요소가 없습니다.");
            using(ZipArchive archive=new ZipArchive(stream,ZipArchiveMode.Read))
            foreach(ZipArchiveEntry entry in archive.Entries)
            {
                string path=Path.GetFullPath(Path.Combine(directory,entry.FullName));
                if(!path.StartsWith(prefix,StringComparison.OrdinalIgnoreCase))throw new IOException("설치 파일의 경로가 올바르지 않습니다.");
                if(String.IsNullOrEmpty(entry.Name)){Directory.CreateDirectory(path);continue;}
                Directory.CreateDirectory(Path.GetDirectoryName(path));
                using(Stream input=entry.Open())using(Stream output=new FileStream(path,FileMode.CreateNew,FileAccess.Write))input.CopyTo(output);
            }
        }
        string[] packages=Directory.GetDirectories(directory);
        if(packages.Length!=1)throw new IOException("설치 파일의 구성 형식이 올바르지 않습니다.");
        VerifyPayload(packages[0]);return packages[0];
    }
    internal static void VerifyPayload(string package)
    {
        string prefix=Path.GetFullPath(package).TrimEnd(Path.DirectorySeparatorChar)+Path.DirectorySeparatorChar;
        string[] lines=File.ReadAllLines(Path.Combine(package,"SHA256SUMS.txt"));
        if(lines.Length<8)throw new IOException("설치 파일의 검증 정보가 없습니다.");
        foreach(string line in lines)
        {
            if(line.Length<67||line.Substring(64,2)!="  ")throw new IOException("설치 파일 검증 형식 오류.");
            string path=Path.GetFullPath(Path.Combine(package,line.Substring(66).Replace('/',Path.DirectorySeparatorChar)));
            if(!path.StartsWith(prefix,StringComparison.OrdinalIgnoreCase))throw new IOException("설치 파일 검증 경로 오류.");
            using(SHA256 hash=SHA256.Create())using(Stream file=File.OpenRead(path))
                if(BitConverter.ToString(hash.ComputeHash(file)).Replace("-","").ToLowerInvariant()!=line.Substring(0,64).ToLowerInvariant())throw new IOException("설치 파일이 손상되었습니다. 다시 다운로드하세요.");
        }
    }
    internal static int Run(bool checkOnly,Action<string> report)
    {
        string directory=Path.Combine(Path.GetTempPath(),"LabPhotoToolsSetup-"+Guid.NewGuid().ToString("N"));
        try
        {
            string error=PrerequisiteError();if(error!=null){report(error);return 2;}
            report("설치 파일을 확인하는 중입니다.");string package=ExtractPayload(directory);
            string shell=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows),Environment.Is64BitProcess?@"System32\WindowsPowerShell\v1.0\powershell.exe":@"Sysnative\WindowsPowerShell\v1.0\powershell.exe");
            ProcessStartInfo start=new ProcessStartInfo(shell,"-NoProfile -NonInteractive -ExecutionPolicy Bypass -File \""+Path.Combine(package,@"scripts\Setup-Entry.ps1")+"\" -SourceDirectory \""+package+"\""+(checkOnly?" -CheckOnly":""));
            start.UseShellExecute=false;start.CreateNoWindow=true;start.RedirectStandardOutput=true;start.RedirectStandardError=true;
            start.StandardOutputEncoding=Encoding.UTF8;start.StandardErrorEncoding=Encoding.UTF8;
            using(Process process=new Process {StartInfo=start})
            {
                process.OutputDataReceived+=delegate(object sender,DataReceivedEventArgs e){if(e.Data!=null)report(e.Data);};
                process.ErrorDataReceived+=delegate(object sender,DataReceivedEventArgs e){if(e.Data!=null)report(e.Data);};
                process.Start();process.BeginOutputReadLine();process.BeginErrorReadLine();process.WaitForExit();return process.ExitCode;
            }
        }
        catch(Exception ex){report(ex.Message);return 1;}
        finally
        {
            // Delete only this invocation's GUID directory, never the install
            // directory or user files. Do not follow unexpected reparse points.
            try
            {
                if(Directory.Exists(directory)&&Path.GetDirectoryName(directory).Equals(Path.GetTempPath().TrimEnd(Path.DirectorySeparatorChar),StringComparison.OrdinalIgnoreCase)
                    && (File.GetAttributes(directory)&FileAttributes.ReparsePoint)==0)
                {
                    bool safe=Directory.GetDirectories(directory,"*",SearchOption.AllDirectories).All(p=>(File.GetAttributes(p)&FileAttributes.ReparsePoint)==0);
                    if(safe)Directory.Delete(directory,true);
                }
            }catch(IOException){}catch(UnauthorizedAccessException){}
        }
    }
}
internal sealed class SetupWindow:Form
{
    private readonly TextBox log=new TextBox {Multiline=true,ReadOnly=true,ScrollBars=ScrollBars.Vertical,Dock=DockStyle.Fill};
    private readonly Label status=new Label {AutoSize=true,Dock=DockStyle.Fill,Text="PowerPoint를 닫은 뒤 ‘설치’를 누르세요."};
    private readonly ProgressBar progress=new ProgressBar {Dock=DockStyle.Fill,Style=ProgressBarStyle.Blocks};
    private readonly Button install=new Button {Text="설치",AutoSize=true,Padding=new Padding(14,5,14,5)};
    private readonly Button close=new Button {Text="닫기",AutoSize=true,Padding=new Padding(14,5,14,5)};
    private readonly Button showLog=new Button {Text="로그 보기",AutoSize=true,Padding=new Padding(8,5,8,5)};
    private readonly Action<string> writeLog;
    private bool running;
    internal SetupWindow(Action<string> writeLog,string logPath)
    {
        this.writeLog=writeLog;Text="Lab Photo Tools "+SetupCore.Version+" 설치";Font=new Font("맑은 고딕",10);AutoScaleMode=AutoScaleMode.Dpi;AutoScaleDimensions=new SizeF(96,96);
        ClientSize=new Size(650,480);MinimumSize=new Size(380,320);StartPosition=FormStartPosition.CenterScreen;
        TableLayoutPanel root=new TableLayoutPanel {Dock=DockStyle.Fill,Padding=new Padding(20),ColumnCount=1,RowCount=6,AutoScroll=true};
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));
        for(int i=0;i<6;i++)root.RowStyles.Add(new RowStyle(i==3?SizeType.Percent:SizeType.AutoSize,i==3?100:0));
        root.Controls.Add(new Label {Text="Lab Photo Tools 설치",Font=new Font(Font,FontStyle.Bold),AutoSize=true,Dock=DockStyle.Fill,Margin=new Padding(0,0,0,12)},0,0);
        root.Controls.Add(new Label {Text="Python이 없어도 설치할 수 있습니다.\n처음 한 번 인터넷으로 Python·사진 처리 도구·모델을 자동 준비합니다.\n다운로드 약 220 MB · 설치 공간 약 1 GB\nWindows 10/11 (Intel/AMD 64비트) 및 데스크톱 PowerPoint 필요\nMicrosoft 구성 요소가 없으면 Windows 관리자 승인이 요청될 수 있습니다.",AutoSize=true,Dock=DockStyle.Fill,Margin=new Padding(0,0,0,14)},0,1);
        root.Controls.Add(status,0,2);root.Controls.Add(log,0,3);root.Controls.Add(progress,0,4);
        FlowLayoutPanel actions=new FlowLayoutPanel {Dock=DockStyle.Fill,AutoSize=true,FlowDirection=FlowDirection.RightToLeft};actions.Controls.Add(close);actions.Controls.Add(install);actions.Controls.Add(showLog);root.Controls.Add(actions,0,5);Controls.Add(root);
        close.Click+=delegate{Close();};showLog.Click+=delegate{Process.Start("notepad.exe","\""+logPath+"\"");};install.Click+=async delegate{await Install();};
        FormClosing+=delegate(object sender,FormClosingEventArgs e){if(running){e.Cancel=true;status.Text="설치가 진행 중입니다. 완료될 때까지 기다려 주세요.";}};
        Shown+=delegate
        {
            Rectangle area=Screen.FromHandle(Handle).WorkingArea;float scale=DeviceDpi/96f;
            // Some .NET Framework hosts report 96 DPI while GDI renders fonts
            // at the display scale. Size the window from the actual font too.
            using(Graphics graphics=CreateGraphics())
            using(Font baseline=new Font("맑은 고딕",10))
                scale=Math.Max(scale,Font.GetHeight(graphics)/baseline.GetHeight(96));
            MinimumSize=new Size(Math.Min((int)(380*scale),area.Width-20),Math.Min((int)(320*scale),area.Height-20));
            Size=new Size(Math.Min((int)(670*scale),area.Width-20),Math.Min((int)(530*scale),area.Height-20));
            Location=new Point(area.X+(area.Width-Width)/2,area.Y+(area.Height-Height)/2);
        };
    }
    private void Report(string line)
    {
        writeLog(line);if(IsDisposed||!IsHandleCreated)return;
        BeginInvoke((MethodInvoker)delegate{if(!IsDisposed){log.AppendText(line+Environment.NewLine);if(line.StartsWith("[1/4]"))status.Text="1/4 · 전용 Python을 준비하는 중";else if(line.StartsWith("[2/4]"))status.Text="2/4 · 사진 처리 라이브러리를 설치하는 중";else if(line.StartsWith("[3/4]"))status.Text="3/4 · 배경 제거 모델을 준비하는 중";else if(line.StartsWith("[4/4]"))status.Text="4/4 · PowerPoint 추가 기능을 등록하는 중";}});
    }
    private async Task Install()
    {
        running=true;install.Enabled=close.Enabled=false;progress.Style=ProgressBarStyle.Marquee;status.Text="설치 환경을 확인하는 중입니다.";
        int result=await Task.Run(()=>SetupCore.Run(false,Report));
        progress.Style=ProgressBarStyle.Blocks;running=false;close.Enabled=true;
        if(result==0){progress.Value=100;status.Text="설치 완료! PowerPoint에서 Lab Photo Tools 탭을 여세요.";install.Text="설치 완료";close.Focus();}
        else{status.Text="설치를 완료하지 못했습니다. 아래 내용을 확인하고 다시 시도하세요.";install.Enabled=true;install.Text="다시 시도";}
    }
}
internal static class SetupProgram
{
    [STAThread] private static int Main(string[] args)
    {
        bool created;using(Mutex mutex=new Mutex(true,"Local\\LabPhotoToolsSetup",out created))
        {
            if(!created){if(!args.Contains("/quiet"))MessageBox.Show("다른 Lab Photo Tools 설치가 진행 중입니다.");return 3;}
            string logs=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"LabPhotoToolsSetup","logs");Directory.CreateDirectory(logs);
            string path=Path.Combine(logs,"setup-"+DateTime.Now.ToString("yyyyMMdd-HHmmss")+"-"+Guid.NewGuid().ToString("N").Substring(0,6)+".log");
            object gate=new object();using(StreamWriter writer=new StreamWriter(path,false,new UTF8Encoding(true)))
            {
                writer.AutoFlush=true;Action<string> report=line=>{lock(gate)writer.WriteLine(line);};report("Lab Photo Tools "+SetupCore.Version+" Setup");
                if(args.Contains("/quiet")||args.Contains("/check"))return SetupCore.Run(args.Contains("/check"),report);
                Application.EnableVisualStyles();Application.SetCompatibleTextRenderingDefault(false);Application.Run(new SetupWindow(report,path));return 0;
            }
        }
    }
}
