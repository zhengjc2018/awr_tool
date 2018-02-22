init_para_cmplist={
"sec_case_sensitive_logon":("FALSE","该参数为大小写敏感参数，11g之后默认开启，但多数应用开发人员并不知道密码大小写敏感，极易引发登录错误问题，建议关闭。"), 
"deferred_segment_creation":("FALSE","该参数为延迟段创建特性，11g之后创建空表默认不立即创建段空间，对于erp等大型应用可以节约一定空间，但是10g exp导出工具无法导出此类空表，建议关闭。"), 
"audit_trail":("NONE","该参数为oracle审计，11g之后默认为db，审计信息存储于system表空间aud$表格，实际使用情况经常将system表空间撑满，建议关闭。"),
"_serial_direct_read":("NEVER","该参数为直接路径读参数，11g之后，oracle会将小表的全表扫描转化成直接路径读（大表小表的界限由_small_table_threshold参数设置），实际情况往往会加大内存压力，导致性能问题，建议关闭。"),
"_use_adaptive_log_file_sync":("FALSE","该参数为自动切换日志写模式，11g之后默认为true。在实际使用中，Oracle会在commit时自动使用“延迟提交”模式，这将导致log file sync时间加长，commit响应时间增大。建议关闭。 "),
"_optimizer_cost_based_transformation":("OFF","该参数为CBO模式下的查询转化，实际使用情况在默写情况下会导致执行计划异常影响性能，建议关闭。"),
"_external_scn_rejection_threshold_hours":("1","根据实际经验，此参数可以防止出现无效的SCN的最佳实践。"),
"_external_scn_logging_threshold_seconds":("600","出现无效SCN后，可通过此参数追踪问题根源。"),
"event":("28401 TRACE NAME CONTEXT FOREVER, LEVEL 1","该参数为11g密码延迟登录新特性，当业务用户密码连续错误之后，oracle会延迟正确密码登录，造成Oracle内部锁竞争（library cache lock和某些Mutex），严重影响业务性能，建议关闭。"),
"parallel_force_local":("TRUE","该参数为本地并行查询，在Oracle RAC多实例环境中，并行操作会被分配到多个节点并发执行，但这将导致极难控制资源消耗、节点间交互过多。建议设置此参数，关闭多节点并行执行单一SQL。"),
"fast_start_parallel_rollback":("HIGH","该参数为数据库打开时的快速并行回滚，在资源充足情况下，建议这设为HIGH，可以减少大事务失败回滚时间，从而增加数据库Open的时间。"),
"_gc_read_mostly_locking":("FALSE","该参数为oracle RAC DRM特性，实际使用中有可能触发竞争，并严重影响性能，建议关闭。"),
"_gc_policy_time":("0","该参数为oracle RAC DRM特性，实际使用中有可能触发竞争、并严重影响性能，建议关闭。"),
"_clusterwide_global_transactions":("FALSE","该参数为集群级别分布式事务，将某一个事务切分到多个节点上，此特性存在bug，建议关闭。")
}