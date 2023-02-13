select
    tbls.relname        as  "table_physical_name"
,   clmn.column_name    as  "column_physical_name"
,   dscr.description    as  "column_logical_name"
from
    pg_stat_user_tables tbls
inner join
    information_schema.columns  clmn
on
    tbls.relname    =   clmn.table_name
left outer join
    pg_description      dscr
on
    dscr.objoid     =   tbls.relid
    and
    dscr.objsubid   =   clmn.ordinal_position
where
    tbls.schemaname =   %s  -- スキーマ名
    and
    tbls.relname    =   %s  -- テーブル名
;