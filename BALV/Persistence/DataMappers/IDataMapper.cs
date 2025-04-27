using System;
using System.Data;

namespace BAL.Persistence.DataMappers
{
    interface IDataMapper
    {
        // Main method that populates dto with data
        Object GetData(IDataReader reader);
        // Gets the num results returned. Needed for data paging.
        int GetRecordCount(IDataReader reader);
    }
}
