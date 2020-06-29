using System;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace NetOfficeSamples.PowerPointAutomation
{
    class Program
    {
        static void Main(string[] args)
        {
            var app = new Application();
            var presentation = app.Presentations.Add(MsoTriState.msoTrue);
            var slide = presentation.Slides.Add(1, PpSlideLayout.ppLayoutText);

            if (slide.Shapes.Count > 1)
            {
                var shape = slide.Shapes[1];
                shape.TextFrame.TextRange.Text = "Welcome to PowerPoint";
            }

            app.Quit();

            Console.WriteLine("Hello World!");
        }
    }
}
