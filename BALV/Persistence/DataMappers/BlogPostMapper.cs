using System;
using System.Data;
using GRN.Common.DataShapes;


// The convention for DataMappers is that their name is of the form <TypeName>Mapper.
// So a mapper for the type BlogPost will be named BlogPostMapper.  

namespace GRN.BAL.Persistence.DataMappers
{
    class BlogPostMapper : IDataMapper
    {
        private bool _isInitialized = false;
        private int _ordinal_PostGuid;
        private int _ordinal_PostId;
        private int _ordinal_CreatedUtc;
        private int _ordinal_ModifiedUtc;
        private int _ordinal_PostedUtc;
        private int _ordinal_BlogGuid;
        private int _ordinal_PostUrl;
        private int _ordinal_PostTitle;
        private int _ordinal_PostSummary;
        private int _ordinal_PostBody;
        private int _ordinal_AuthorName;
        private int _ordinal_Score;


        private void InitializeMapper(IDataReader reader)
        {
            PopulateOrdinals(reader);
            _isInitialized = true;
        }


        public void PopulateOrdinals(IDataReader reader)
        {
            _ordinal_PostGuid = reader.GetOrdinal("PostGuid");
            _ordinal_PostId = reader.GetOrdinal("PostId");
            _ordinal_CreatedUtc = reader.GetOrdinal("CreatedUtc");
            _ordinal_ModifiedUtc = reader.GetOrdinal("ModifiedUtc");
            _ordinal_PostedUtc = reader.GetOrdinal("PostedUtc");
            _ordinal_BlogGuid = reader.GetOrdinal("BlogGuid");
            _ordinal_PostUrl = reader.GetOrdinal("PostUrl");
            _ordinal_PostTitle = reader.GetOrdinal("PostTitle");
            _ordinal_PostSummary = reader.GetOrdinal("PostSummary");
            _ordinal_PostBody = reader.GetOrdinal("PostBody");
            _ordinal_AuthorName = reader.GetOrdinal("AuthorName");
            _ordinal_Score = reader.GetOrdinal("Score");
        }


        public Object GetData(IDataReader reader)
        {
            // This is where we define the mapping between the object properties and the 
            // data columns. The convention that should be used is that the object property 
            // names are exactly the same as the column names. However if there is some 
            // compelling reason for the names to be different the mapping can be defined here.

            // We assume the reader has data and is already on the row that contains the data 
            //we need. We don't need to call read. As a general rule, assume that every field must 
            //be null  checked. If a field is null then the nullvalue for that  field has already 
            //been set by the DTO constructor, we don't have to change it.

            if (!_isInitialized) { InitializeMapper(reader); }
            BlogPost dto = new BlogPost();
            // Initialize score to 0. This is the default
            // we want to use if there is no score.
            dto.Score = 0;
            // Now we can load the data
            if (!reader.IsDBNull(_ordinal_PostGuid)) { dto.PostGuid = reader.GetGuid(_ordinal_PostGuid); }
            if (!reader.IsDBNull(_ordinal_PostId)) { dto.PostId = reader.GetInt32(_ordinal_PostId); }
            if (!reader.IsDBNull(_ordinal_CreatedUtc)) { dto.CreatedUtc = reader.GetDateTime(_ordinal_CreatedUtc); }
            if (!reader.IsDBNull(_ordinal_ModifiedUtc)) { dto.ModifiedUtc = reader.GetDateTime(_ordinal_ModifiedUtc); }
            if (!reader.IsDBNull(_ordinal_PostedUtc)) { dto.PostedUtc = reader.GetDateTime(_ordinal_PostedUtc); }
            if (!reader.IsDBNull(_ordinal_BlogGuid)) { dto.BlogGuid = reader.GetGuid(_ordinal_BlogGuid); }
            if (!reader.IsDBNull(_ordinal_PostUrl)) { dto.PostUrl = reader.GetString(_ordinal_PostUrl); }
            if (!reader.IsDBNull(_ordinal_PostTitle)) { dto.PostTitle = reader.GetString(_ordinal_PostTitle); }
            if (!reader.IsDBNull(_ordinal_PostSummary)) { dto.PostSummary = reader.GetString(_ordinal_PostSummary); }
            if (!reader.IsDBNull(_ordinal_PostBody)) { dto.PostBody = reader.GetString(_ordinal_PostBody); }
            if (!reader.IsDBNull(_ordinal_AuthorName)) { dto.AuthorName = reader.GetString(_ordinal_AuthorName); }
            if (!reader.IsDBNull(_ordinal_Score)) { dto.Score = reader.GetInt32(_ordinal_Score); }
            return dto;
        }


        public int GetRecordCount(IDataReader reader)
        {
            Object count = reader["RecordCount"];
            return count==null ? 0 : Convert.ToInt32(count) ;
        }
    }
}
