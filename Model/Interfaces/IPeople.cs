﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MacValvesWordGenerate.Model.Interfaces
{
    public interface IPeople
    {
        string Name { get; set; }
        string Surname { get; set; }
        string Function { get; set; }
        string Customer { get; set; }
    }
}