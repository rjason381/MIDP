from __future__ import annotations
import json
import sys
from pathlib import Path

if len(sys.argv) != 4:
    print("Uso: python merge_midp_project_into_state.py <backup_actual.json> <project_only.json> <salida.json>")
    raise SystemExit(1)

state_path = Path(sys.argv[1])
project_path = Path(sys.argv[2])
out_path = Path(sys.argv[3])

state = json.loads(state_path.read_text(encoding="utf-8"))
project = json.loads(project_path.read_text(encoding="utf-8"))
projects = state.get("projects")
if not isinstance(projects, list):
    raise SystemExit("El backup no tiene 'projects' valido.")
project_id = project.get("id")
projects = [item for item in projects if not (isinstance(item, dict) and item.get("id") == project_id)]
projects.append(project)
state["projects"] = projects
state["activeProjectId"] = project_id
out_path.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
print(f"Estado combinado guardado en: {out_path}")
