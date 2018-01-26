using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JTools.MLO
{
    public class MLOTask
    {
        private MLOTaskGeneral _MLOTaskGeneral = new MLOTaskGeneral();

        public string Title { get; set; }

        public string Note { get; set; } 

        public MLOTaskGeneral General
        {
            get { return _MLOTaskGeneral; }
            set { _MLOTaskGeneral = value;  }
        }
    }
}
