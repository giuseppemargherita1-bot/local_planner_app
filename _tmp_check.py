import sqlite3
conn = sqlite3.connect(r"C:\Users\UTENTE\Desktop\PROJ\local_planner_app\planner.db")
conn.row_factory = sqlite3.Row
print('PROJECTS')
for r in conn.execute("select id, code, type, workshop_rollup from projects where code like '%328%' or code='OVERALL OFFICINA' order by code"):
    print(dict(r))
print('DEMANDS')
for r in conn.execute("select p.code, d.role, d.week, d.qty from demands d join projects p on p.id=d.project_id where p.code like '%328%' and d.week between 13 and 15 order by p.code, d.role, d.week"):
    print(dict(r))
print('ALLOC')
for r in conn.execute("select p.code, a.role, a.week_from, a.week_to, rs.name from allocations a join projects p on p.id=a.project_id join resources rs on rs.id=a.resource_id where p.code like '%328%' and a.week_to>=13 and a.week_from<=15 order by p.code, a.role, rs.name"):
    print(dict(r))
