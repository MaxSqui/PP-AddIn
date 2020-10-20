﻿using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointAddInVSTO.Extensions
{
    public static class PresentationExtensions
    {
        public static IEnumerable<string> GetMediaNames(this Presentation presentation)
        {
            foreach (Slide slide in presentation.Slides)
            {
                foreach (Shape shape in slide.Shapes)
                {
                    if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoMedia)
                    {
                        yield return shape.Name;
                    }
                }
            }
        }

        public static IEnumerable<Shape> GetAnimateShapes(this Presentation presentation)
        {
            foreach (Slide slide in presentation.Slides)
            {
                foreach (Shape shape in slide.Shapes)
                {
                    if (shape.AnimationSettings.Animate == Microsoft.Office.Core.MsoTriState.msoTrue &&
                        shape.Type != Microsoft.Office.Core.MsoShapeType.msoMedia)
                    {
                        yield return shape;
                    }
                }
            }
        }
    }
}