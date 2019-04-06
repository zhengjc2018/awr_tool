{%set err_flag='/' -%}
{%set err_info='原生报告无此项内容' -%}
{%set BackgroundWaitEvent=kw['BackgroundWaitEvent']() -%}
{%set BufferPoolAdv=kw['BufferPoolAdv']() -%}
{%set BufferPoolStat=kw['BufferPoolStat']() -%}
{%set DbInstInfo=kw['DbInstInfo']() -%}
{%set DynamicRemastering=kw['DynamicRemastering']() -%}
{%set FileIOStat=kw['FileIOStat']() -%}
{%set ForegroundWaitEvent=kw['ForegroundWaitEvent']() -%}
{%set HostInfo=kw['HostInfo']() -%}
{%set InitOraParam=kw['InitOraParam']() -%}
{%set InstActivityStat=kw['InstActivityStat']() -%}
{%set InstEfficiencyPer=kw['InstEfficiencyPer']() -%}
{%set IoProfile=kw['IoProfile']() -%}
{%set GCEnqueueService=kw['GCEnqueueService']() -%}
{%set GlobalCacheLoadProfile=kw['GlobalCacheLoadProfile']() -%}
{%set KeyInstActivityStat=kw['KeyInstActivityStat']() -%}
{%set LoadProfile=kw['LoadProfile']() -%}
{%set MemDynamicStat=kw['MemDynamicStat']() -%}
{%set MemoryStat=kw['MemoryStat']() -%}
{%set PdbInfo=kw['PdbInfo']() -%}
{%set PgaMemAdv=kw['PgaMemAdv']() -%}
{%set PGATarget=kw['PGATarget']() -%}
{%set PingLatencyStats=kw['PingLatencyStats']() -%}
{%set SegmentsBufferBusy=kw['SegmentsBufferBusy']() -%}
{%set SegmentsDirectReads=kw['SegmentsDirectReads']() -%}
{%set SegmentsLogReads=kw['SegmentsLogReads']() -%}
{%set SegmentsPhyReads=kw['SegmentsPhyReads']() -%}
{%set SegRowLockWaits=kw['SegRowLockWaits']() -%}
{%set SgaTargetAdv=kw['SgaTargetAdv']() -%}
{%set SharePoolAdv=kw['SharePoolAdv']() -%}
{%set SnapshotInfo=kw['SnapshotInfo']() -%}
{%set SqlCpuTime=kw['SqlCpuTime']() -%}
{%set SqlElapsedTime=kw['SqlElapsedTime']() -%}
{%set SqlExecutions=kw['SqlExecutions']() -%}
{%set SqlGets=kw['SqlGets']() -%}
{%set SqlParseCall=kw['SqlParseCall']() -%}
{%set SqlPhyReads=kw['SqlPhyReads']() -%}
{%set SqlVersionCount=kw['SqlVersionCount']() -%}
{%set SystemStatTime=kw['SystemStatTime']() -%}
{%set TimeModelStat=kw['TimeModelStat']() -%}
{%set TotalWaitTime=kw['TotalWaitTime']() -%}
{%set int=CommonMethod.int -%}
{%set re_sub=CommonMethod.re_sub_unneed -%}
{%set get_fix_info=CommonMethod.merge_info_from_tables -%}

# {{ host.name }} Oracle 数据库AWR报告解读

文档日期：{{ date.year }}年{{ date.month }}月{{  date.day }}日

**目录**
