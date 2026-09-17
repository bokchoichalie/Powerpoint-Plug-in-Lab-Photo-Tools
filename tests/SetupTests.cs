using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing;
internal static class SetupTests
{
    [STAThread] private static int Main(string[] args)
    {
        string directory=Path.Combine(Path.GetTempPath(),"LabPhotoTools-SetupTest-"+Guid.NewGuid().ToString("N"));
        try
        {
            string payload=SetupCore.ExtractPayload(directory);
            if(!File.Exists(Path.Combine(payload,"scripts","Runtime-Downloads.ps1")))throw new Exception("Missing private Python bootstrap");
            if(!File.Exists(Path.Combine(payload,"scripts","Setup-Entry.ps1")))throw new Exception("Missing installer entry");
            string file=Path.Combine(payload,"README.md"),original=File.ReadAllText(file);File.AppendAllText(file,"tamper");
            bool rejected=false;try{SetupCore.VerifyPayload(payload);}catch(IOException){rejected=true;}
            if(!rejected)throw new Exception("Corrupt payload accepted");
            string prerequisite=SetupCore.PrerequisiteError();
            if(System.Diagnostics.Process.GetProcessesByName("POWERPNT").Length>0)
            {if(prerequisite==null||!prerequisite.Contains("PowerPoint 창을 닫은"))throw new Exception("Running PowerPoint must block installation");}
            else if(prerequisite!=null)throw new Exception(prerequisite);
            Application.EnableVisualStyles();
            using(SetupWindow form=new SetupWindow(s=>{},file))
            {
                form.Show();Application.DoEvents();
                Console.WriteLine("Setup client="+form.ClientSize+" DPI="+form.DeviceDpi+" font="+form.Font.Size);
                using(Bitmap image=new Bitmap(form.Width,form.Height)){form.DrawToBitmap(image,new Rectangle(Point.Empty,image.Size));image.Save(args[0]);}
                Button install=Descendants(form).OfType<Button>().Single(b=>b.Text=="설치");
                if(!form.ClientRectangle.Contains(form.RectangleToClient(install.RectangleToScreen(install.ClientRectangle))))throw new Exception("Install button is clipped");
                TableLayoutPanel layout=Descendants(form).OfType<TableLayoutPanel>().Single();
                Control[] controls=layout.Controls.Cast<Control>().ToArray();
                foreach(Control c in controls)if(!form.ClientRectangle.Contains(form.RectangleToClient(c.RectangleToScreen(c.ClientRectangle))))throw new Exception("Installer control is clipped: "+c.Text+" "+c.Bounds+" screen="+form.RectangleToClient(c.RectangleToScreen(c.ClientRectangle)));
                for(int i=0;i<controls.Length;i++)for(int j=i+1;j<controls.Length;j++)if(controls[i].Bounds.IntersectsWith(controls[j].Bounds))throw new Exception("Installer controls overlap");
                using(Bitmap image=new Bitmap(form.Width,form.Height)){form.DrawToBitmap(image,new Rectangle(Point.Empty,image.Size));image.Save(args[0]);}
                form.Close();
            }
            Console.WriteLine("PASS: installer payload, corruption rejection, prerequisites and native install window");return 0;
        }
        catch(Exception ex){Console.Error.WriteLine(ex);return 1;}
        finally
        {
            string expected=Path.GetTempPath().TrimEnd(Path.DirectorySeparatorChar);
            if(Directory.Exists(directory)&&Path.GetDirectoryName(directory)==expected&&Path.GetFileName(directory).StartsWith("LabPhotoTools-SetupTest-"))Directory.Delete(directory,true);
        }
    }
    private static System.Collections.Generic.IEnumerable<Control> Descendants(Control root)
    {foreach(Control c in root.Controls){yield return c;foreach(Control child in Descendants(c))yield return child;}}
}
