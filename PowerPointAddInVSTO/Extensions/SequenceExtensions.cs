﻿using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointAddInVSTO.Extensions
{
    public static class SequenceExtensions
    {
        public static IEnumerable<Effect> GetDependentEffects(this Sequence sequence, Effect effect)
        {
            bool inDependentRange = false;
            foreach (Effect dependentEffect in sequence)
            {
                if (dependentEffect == effect) inDependentRange = true;
                if(inDependentRange && dependentEffect.Timing.TriggerType != MsoAnimTriggerType.msoAnimTriggerOnPageClick)
                {
                    yield return dependentEffect;
                }
                if (dependentEffect.Timing.TriggerType == MsoAnimTriggerType.msoAnimTriggerOnPageClick &&  
                    dependentEffect != effect) inDependentRange = false;
            }
        }
    }
}
