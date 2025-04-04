using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EmailGenerator.Interfaces
{
    public interface IFedExAuthProvider
    {
        Task<string> GetAccessTokenAsync();
    }

}
