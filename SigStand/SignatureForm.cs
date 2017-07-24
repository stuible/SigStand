using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;

namespace SigStand
{
    public class SignatureForm
    {
        public SignatureForm()
        {
        }
        public string name { get; set; }
        public string Name { get; internal set; }
        public string title { get; set; }
        public string Title { get; internal set; }
        public string department { get; set; }



    }
}
