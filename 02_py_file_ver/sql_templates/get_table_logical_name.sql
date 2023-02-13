select
    tbls.relname        as  "table_physical_name"
,   dscr.description    as  "table_logical_name"
from
	 pg_stat_user_tables tbls
left outer join     
    pg_description dscr
on
    tbls.relid       =   dscr.objoid
where
    tbls.schemaname  =   %s    -- 任意の値
	and
    (
        dscr.objsubid     =   0
        or
        dscr.objoid       is  null
    )
	and
		tbls.relname     =   %s  -- 任意の値
;