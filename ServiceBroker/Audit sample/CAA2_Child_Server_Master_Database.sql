USE master

------------------------------------------------------------------------------
-- SET UP TRANSPORT SECURITY
------------------------------------------------------------------------------
IF EXISTS(SELECT * FROM sys.endpoints WHERE NAME = 'ServiceBrokerEndPoint')
	DROP ENDPOINT ServiceBrokerEndPoint

IF EXISTS(SELECT * FROM sys.certificates WHERE NAME = 'CertificateAuditDataReceiver')
	DROP CERTIFICATE CertificateAuditDataReceiver

IF EXISTS(SELECT * FROM sys.certificates WHERE NAME = 'CertificateAuditDataSender')
	DROP CERTIFICATE CertificateAuditDataSender

IF EXISTS(SELECT * FROM sys.server_principals WHERE NAME = 'LoginAuditDataSender')
	DROP LOGIN LoginAuditDataSender

IF EXISTS(SELECT * FROM sys.sysusers WHERE NAME = 'UserAuditDataSender')
	DROP USER UserAuditDataSender

IF EXISTS(SELECT * FROM sys.symmetric_keys WHERE NAME = '##MS_DatabaseMasterKey##')
	DROP MASTER KEY

-- create the login that will be used to send the audited data through the Endpoint
CREATE LOGIN LoginAuditDataSender WITH PASSWORD = 'Login_Audit_DataSender_Password'
GO

-- Create a user for our login
CREATE USER UserAuditDataSender FOR LOGIN LoginAuditDataSender
GO

-- create a master key the for master database
CREATE MASTER KEY ENCRYPTION BY PASSWORD = 'Put_Your_Custom_Password_For_Senders_Master_DB_Here'

GO
-- create certificate for the service broker TCP endpoint for secure communication
-- between servers
CREATE CERTIFICATE CertificateAuditDataSender
WITH 
	-- BOL: The term subject refers to a field in the metadata of 
	--		the certificate as defined in the X.509 standard
	SUBJECT = 'CertAuditDataSender',
	-- set the start date
	START_DATE = '01/01/2007', 
	-- set the expiry data
    EXPIRY_DATE = '01/01/2010' 
	-- enables the certifiacte for service broker initiator
	ACTIVE FOR BEGIN_DIALOG = ON

GO
-- copy the certificate create on the master target server 
-- to c:\ disk and recreate it here with the data sender user authorization
CREATE CERTIFICATE CertificateAuditDataReceiver
	AUTHORIZATION UserAuditDataSender
	FROM FILE = 'c:\CertificateAuditDataReceiver.cer'

GO

-- create endpoint which will be used to send audited data to the 
-- MasterAuditServer
CREATE ENDPOINT ServiceBrokerEndPoint
	-- set endpoint to activly listen for connections
	STATE = STARTED
	-- set it for TCP traffic only since service broker supports only TCP protocol
	-- by convention, 4022 is used but any number between 1024 and 32767 is valid.
	AS TCP (LISTENER_PORT = 4022)
	FOR SERVICE_BROKER 
	(
		-- authenticate connections with our certificate
		AUTHENTICATION = CERTIFICATE CertificateAuditDataSender,
		-- default is REQUIRED encryption but let's just set it to SUPPORTED
		-- SUPPORTED means that the data is encrypted only if the 
		-- opposite endpoint specifies either SUPPORTED or REQUIRED.
		ENCRYPTION = SUPPORTED
	)
GO
-- finally grant the connect permissions to login for the endpoint
GRANT CONNECT ON ENDPOINT::ServiceBrokerEndPoint TO LoginAuditDataSender