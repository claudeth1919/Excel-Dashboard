using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Dashboard
{
    public class Header
    {
        private string name;
        private string dataType;
        private int index;

        public Header(string name, int index)
        {
            this.name = name;
            this.index = index;
        }

        public string Name
        {
            get { return name; }
            set { name = value; }
        }

        public string DataType
        {
            get { return dataType; }
            set { dataType = value; }
        }

        public int Index
        {
            get { return index; }
            set { index = value; }
        }
        public override string ToString()
        {
            return $"(Col {index}) {name}";
        }
    }
}
