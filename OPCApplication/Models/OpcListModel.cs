using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OPCApplication.Models
{
    public class OpcListModel
    {
        private string name;
        private string address;
        private string value;
        private string info;

        public string Name
        {
            get
            {
                return name;
            }

            set
            {
                name = value;
            }
        }

        public string Address
        {
            get
            {
                return address;
            }

            set
            {
                address = value;
            }
        }

        public string Value
        {
            get
            {
                return value;
            }

            set
            {
                this.value = value;
            }
        }

        public string Info
        {
            get
            {
                return info;
            }

            set
            {
                info = value;
            }
        }
    }
}
