--CREATE FULLTEXT CATALOG TitlesFullTextCatalog

CREATE FULLTEXT INDEX ON tProduct (P_TITLE)
     KEY INDEX PK_tProduct
          ON TitlesFullTextCatalog
     WITH CHANGE_TRACKING AUTO
