USE master
------------------------------------------------------------------------------
-- SET UP TRANSPORT SECURITY
------------------------------------------------------------------------------

-- drop receiver endpoint
IF EXISTS(SELECT * FROM sys.endpoints WHERE NAME = N'ServiceBrokerEndPoint')
	DROP ENDPOINT ServiceBrokerEndPoint

-- drop existing receiver certificate
IF EXISTS(SELECT * FROM sys.certificates WHERE NAME = N'CertificateAuditDataReceiver')
	DROP CERTIFICATE CertificateAuditDataReceiver

-- drop master key
IF EXISTS(SELECT * FROM sys.symmetric_keys WHERE NAME = N'##MS_DatabaseMasterKey##')
	DROP MASTER KEY

GO
-- create a master key the for master database
CREATE MASTER KEY ENCRYPTION BY PASSWORD = 'Put_Your_Custom_Password_For_Receivers_Master_DB_Here'

GO
-- create certificate for the service broker TCP endpoint for secure communication
-- between servers
CREATE CERTIFICATE CertificateAuditDataReceiver
WITH 
	-- BOL: The term subject refers to a field in the metadata of 
	--		the certificate as defined in the X.509 standard
	SUBJECT = 'CertAuditDataReceiver',
	-- set the start date
	START_DATE = '01/01/2007', 
	-- set the expiry data
    EXPIRY_DATE = '01/01/2010' 
	-- enables the certifiacte for service broker initiator
	ACTIVE FOR BEGIN_DIALOG = ON

GO
-- save certificate to a file
-- copy this file to all servers whose databases will be audited so we can restore it there
-- and thus enable the secure connection
BACKUP CERTIFICATE CertificateAuditDataReceiver 
	TO FILE = 'c:\CertificateAuditDataReceiver.cer'

GO
-- create endpoint which will be used to send audited data to the 
-- MasterAuditServer ServiceBrokerEndPoint
CREATE ENDPOINT ServiceBrokerEndPoint
	-- set endpoint to activly listen for connections
	STATE = STARTED
	-- set it for TCP traffic only since service broker supports only TCP protocol
	-- by convention, 4022 is used but any number between 1024 and 32767 is valid.
	AS TCP (LISTENER_PORT = 4022)
	FOR SERVICE_BROKER 
	(
		-- authenticate connections with our certificate
		AUTHENTICATION = CERTIFICATE CertificateAuditDataReceiver, 
		-- default is REQUIRED encryption but let's just set it to SUPPORTED
		-- SUPPORTED means that the data is encrypted only if the 
		-- opposite endpoint specifies either SUPPORTED or REQUIRED.
		ENCRYPTION = SUPPORTED
	)

GO
-- finally grant the connect permissions to public
GRANT CONNECT ON ENDPOINT::ServiceBrokerEndPoint TO PUBLIC
