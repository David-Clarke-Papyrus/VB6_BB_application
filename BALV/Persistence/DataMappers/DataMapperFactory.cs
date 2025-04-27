using System;
namespace BAL.Persistence.DataMappers
{
    class DataMapperFactory
    {
        public IDataMapper GetMapper(Type dtoType)
        {
            switch(dtoType.Name)
            {
                case "BlogPost":
                   // return new BlogPostMapper();
                default:
                    return new GenericMapper(dtoType);
            }       
        }

    }
}
