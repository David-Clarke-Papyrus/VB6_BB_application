
DECLARE @SecondsFromMidnight INT
DECLARE @JobScheduleId INT, 
	@ScheduledJobId INT, 
	@validFrom DATETIME, 
	@ScheduledJobStepId INT, 
	@secondsOffset INT, 
	@NextRunOn DATETIME
-- run the job daily
SELECT @SecondsFromMidnight = 8000
SELECT    @validFrom = GETUTCDATE(), -- the job is valid from current UTC time
         -- run the job 2 minutes after the validFrom time. 
         -- we need the offset in seconds from midnight of that day for all jobs
          @secondsOffset = @SecondsFromMidnight, 
          @NextRunOn = DATEADD(n, 1, @validFrom) 
EXEC usp_AddJobSchedule @JobScheduleId OUT,
                        @RunAtInSecondsFromMidnight = @secondsOffset,
                        @FrequencyType = 1,
                        @Frequency = 1 -- run every day                      
-- add new scheduled job 
EXEC usp_AddScheduledJob @ScheduledJobId OUT, @JobScheduleId, 'Papyrus nightly updates', @validFrom
DECLARE @SQL NVARCHAR(MAX)
SELECT  @SQL = N'USE PBKS;   DECLARE @D DATETIME;  SELECT @D = GETDATE();  EXEC dbo.STATS2 @D;EXEC dbo.CreateStatsSet;EXEC sp_DAYEND2'
-- EXEC sp_DAYEND2
EXEC usp_AddScheduledJobStep @ScheduledJobStepId OUT, @ScheduledJobId, @SQL, 'step 1'
-- start the scheduled job
EXEC usp_StartScheduledJob @ScheduledJobId 
