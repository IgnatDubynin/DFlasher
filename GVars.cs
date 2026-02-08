using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DFlasher
{
    public struct Stimulus
    {
        public int Digits { get; set; }
        public int Frequency { get; set; }
        public int Duration { get; set; }
        public bool NeedReaction { get; set; }

        public Stimulus(bool AnyFlag = true)//прикол в том, что если при создании экземпляра умолчать флаг, то конструктор не вызывается, поэтому если надо, пихаем любое значение флага
        {
            Digits = 55;
            Frequency = 40;
            Duration = 5000;
            NeedReaction = true;
        }
    }

    class GVars
    {
        public static bool LockCtrl = false;
    }
}
