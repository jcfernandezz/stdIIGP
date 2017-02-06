using System;
using System.Data.Entity;
using System.Data.Entity.Infrastructure;

namespace IntegradorDeGP
{
    public partial class EntitiesGP : DbContext
    {
        public EntitiesGP(String connectionString): base(connectionString)
        {

        }
    }
}
