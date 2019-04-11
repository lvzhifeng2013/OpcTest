using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpcTest
{
    class Test
    {
        private string id;
        private string value;
        private string description;
        private string content;

       
        public string Id
        {
            get
            {
                return id;
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
        public string Description
        {
            get
            {
                return description;
            }
        }

        public string Content
        {
            get
            {
                return content;
            }

            set
            {
                content = value;
            }
        }

        public Test(string id,string content, string description)
        {
            this.id = id;
            this.description = description;
            this.content = content;
        }

    }
}
