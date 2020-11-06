﻿using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace PowerPointAddInVSTO.ViewModel
{
    public class EffectViewModel : INotifyPropertyChanged,  IDataErrorInfo
    {

        public event PropertyChangedEventHandler PropertyChanged;
        public void OnPropertyChanged([CallerMemberName] string prop = "")
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(prop));
        }

        public int Id { get; set; }
        public string DisplayName { get; set; }
        public Slide Slide { get; set; }
        public int SlideNumber { get; set; }
        public MsoShapeType Type { get; set; }
        private float effectTimeline { get; set; }

        public float LastEffectTimeline { get; set; }
        public int LastSlideNumber { get; set; }
        public float EffectTimeline 
        { 
            get { return effectTimeline; } 
            set 
            {
                effectTimeline = value;
                OnPropertyChanged("EffectTimeline");
            } 
        }

        public bool IsSec { get; set; }

        public string Error
        {
            get { throw new NotImplementedException(); }
        }

        public string this[string columnName]
        {
            get
            {
                string error = String.Empty;
                switch (columnName)
                {
                    case "EffectTimeline":
                        if (EffectTimeline < LastEffectTimeline && SlideNumber == LastSlideNumber)
                        {
                            error = "Value can not be less than previous";
                        }
                        break; 
                }
                return error;
            }
        }

        public Effect FindEffectById(int Id, IEnumerable<Effect> effects)
        {
            foreach (var effect in effects)
            {
                if (effect.Index == Id) return effect;
            }
            return null;
        }

    }
}