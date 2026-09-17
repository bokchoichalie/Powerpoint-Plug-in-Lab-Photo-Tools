using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using LabPhotoTools;

// A short, real-PowerPoint acceptance check for the most recent layout and
// photo-label workflow. It deliberately has no Python/image-processing work.
internal static class QuickPowerPointValidation
{
    private static int checks;
    private static void Check(bool condition, string message)
    {
        checks++;
        if (!condition) throw new Exception(message);
    }
    [STAThread]
    private static int Main(string[] args)
    {
        dynamic app = null, presentation = null;
        try
        {
            string output = args[0]; Directory.CreateDirectory(output);
            string image = Path.Combine(output, "quick-validation.png");
            using (Bitmap bitmap = new Bitmap(120, 80)) using (Graphics graphics = Graphics.FromImage(bitmap))
            { graphics.Clear(Color.LightSteelBlue); graphics.FillEllipse(Brushes.SteelBlue, 25, 10, 70, 60); bitmap.Save(image); }
            app = Activator.CreateInstance(Type.GetTypeFromProgID("PowerPoint.Application"));
            presentation = app.Presentations.Add(-1);
            presentation.PageSetup.SlideWidth = 960f; presentation.PageSetup.SlideHeight = 540f;
            PowerPointHost host = new PowerPointHost(app);

            dynamic arranged = presentation.Slides.Add(1, 12); app.ActiveWindow.View.GotoSlide(1);
            for (int i = 0; i < 11; i++) arranged.Shapes.AddPicture(image, 0, -1, 15f + i * 25f, 40f + i * 12f, 100f, 70f);
            try { app.ActiveWindow.Selection.Unselect(); } catch { }
            host.ArrangeAllPhotos();
            SelectionSnapshot plan = host.ReadAllPhotos();
            Check(plan.Photos.GroupBy(photo => photo.Top).Select(row => row.Count()).SequenceEqual(new[] { 4, 4, 3 }), "11 photographs must use 4-4-3");

            dynamic labels = presentation.Slides.Add(2, 12); app.ActiveWindow.View.GotoSlide(2);
            dynamic lowerRight = labels.Shapes.AddPicture(image, 0, -1, 270f, 230f, 120f, 80f);
            dynamic upperRight = labels.Shapes.AddPicture(image, 0, -1, 260f, 40f, 120f, 80f);
            dynamic upperLeft = labels.Shapes.AddPicture(image, 0, -1, 45f, 48f, 120f, 80f);
            dynamic square = host.AddNumberLabel("square", 0);
            dynamic circle = host.AddNumberLabel("circle", 0);
            dynamic alphabet = host.AddNumberLabel("paren", 0);
            Check((string)square.TextFrame2.TextRange.Text == "0" && (string)alphabet.TextFrame2.TextRange.Text == "(A)", "Preset text");
            CheckPair(upperLeft, square, "upper-left"); CheckPair(upperRight, circle, "upper-right"); CheckPair(lowerRight, alphabet, "lower-right");
            double leftBeforeArrange = (double)upperLeft.Left;
            try { app.ActiveWindow.Selection.Unselect(); } catch { }
            host.ArrangeAllPhotos();
            CheckPair(upperLeft, square, "upper-left after smart arrange");
            CheckPair(upperRight, circle, "upper-right after smart arrange");
            CheckPair(lowerRight, alphabet, "lower-right after smart arrange");
            Check(Math.Abs((double)upperLeft.Left - leftBeforeArrange) > .05, "Smart arrange must move the labelled photograph group");
            int relabelPhotoId = (int)upperRight.Id;
            circle.Delete();
            dynamic replacement = host.AddNumberLabel("circle", 1);
            dynamic relabelledPhoto = PhotoInLabelGroup(replacement);
            Check((int)relabelledPhoto.Id == relabelPhotoId && (string)replacement.TextFrame2.TextRange.Text == "1", "Deleted label photograph must be selected again");
            CheckPair(relabelledPhoto, replacement, "relabelled photograph");
            dynamic standalone = host.AddNumberLabel("suffix", 0);
            Check((string)standalone.TextFrame2.TextRange.Text == "0)" && (int)labels.Shapes.Count == 4, "Standalone label when all photos are labelled");

            string saved = Path.Combine(output, "LabPhotoTools-0.1.17-quick-validation.pptx");
            presentation.SaveAs(saved, 24); presentation.Close(); presentation = null;
            presentation = app.Presentations.Open(saved, 0, 0, -1); app.ActiveWindow.View.GotoSlide(2);
            dynamic added = presentation.Slides.Item(2).Shapes.AddPicture(image, 0, -1, 500f, 45f, 100f, 70f);
            dynamic next = host.AddNumberLabel("square", null);
            Check((string)next.TextFrame2.TextRange.Text == "11", "Number sequence persists after reopening");
            CheckPair(added, next, "reopened photo");
            dynamic empty = presentation.Slides.Add(3, 12); app.ActiveWindow.View.GotoSlide(3);
            dynamic emptyLabel = host.AddNumberLabel("circle", 5);
            Check((string)emptyLabel.TextFrame2.TextRange.Text == "5" && (int)empty.Shapes.Count == 1, "Standalone label on a slide with no photo");
            Console.WriteLine("PASS: " + checks + " quick PowerPoint acceptance assertions.");
            return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
        finally
        {
            if (presentation != null) { try { presentation.Saved = -1; presentation.Close(); } catch { } }
            if (app != null) { try { app.Quit(); } catch { } try { Marshal.ReleaseComObject(app); } catch { } }
        }
    }
    private static void CheckPair(dynamic photo, dynamic label, string name)
    {
        Check((string)photo.Tags.Item("LABPHOTO_ATTACHED_LABEL") == "1", name + " photo tag");
        dynamic group = photo.ParentGroup;
        Check((int)group.Type == 6 && (int)group.GroupItems.Count == 2, name + " group");
        Check(Math.Abs((double)photo.Left - (double)label.Left) < .05 && Math.Abs((double)photo.Top - (double)label.Top) < .05, name + " top-left alignment");
    }
    private static dynamic PhotoInLabelGroup(dynamic label)
    {
        dynamic group = label.ParentGroup;
        for (int i = 1; i <= (int)group.GroupItems.Count; i++)
        {
            dynamic item = group.GroupItems.Item(i);
            int type = (int)item.Type;
            if (type == 13 || type == 11) return item;
            if (type == 14)
            {
                try { int contained = (int)item.PlaceholderFormat.ContainedType; if (contained == 13 || contained == 11) return item; }
                catch { }
            }
        }
        throw new Exception("Label group has no photograph");
    }
}
