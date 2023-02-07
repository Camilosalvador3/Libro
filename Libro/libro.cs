using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Libro
{
    internal class libro
    {
        double ISBN, NùmeroDePag;
        string Tìtulo, Autor;
        #region "constructor"
        public libro(double iSBN, double nùmeroDePag, string tìtulo, string autor)
        {
            ISBN1 = iSBN;
            NùmeroDePag1 = nùmeroDePag;
            Tìtulo1 = tìtulo;
            Autor1 = autor;
        }
        #endregion
        #region "get and setter"
        public double ISBN1 
        { 
          get => ISBN; 
          set => ISBN = value; 
        }
        public double NùmeroDePag1 
        { 
          get => NùmeroDePag; 
          set => NùmeroDePag = value; 
        }
        public string Tìtulo1 
        { 
          get => Tìtulo; 
          set => Tìtulo = value; 
        }
        public string Autor1 
        { 
          get => Autor; 
          set => Autor = value; 
        }
        #endregion

        #region "Metodo Mostrar"
        public class Mostrar
        {

        }
        #endregion
    }
}
