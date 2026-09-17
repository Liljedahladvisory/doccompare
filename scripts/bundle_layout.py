"""Keep source overlays consistent with py2app's relocated subpackages."""
import ast
from pathlib import Path
import shutil


def refresh_relocated_packages(source: Path, lib_dir: Path, boot: Path):
    hooks = []
    for node in ast.parse(boot.read_text()).body:
        if isinstance(node, ast.Assign) and any(
            isinstance(target, ast.Name) and target.id == '_path_hooks'
            for target in node.targets
        ):
            hooks = ast.literal_eval(node.value)
    refreshed = []
    for name in hooks:
        if not name.startswith('doccompare.'):
            continue
        parts = name.split('.')[1:]
        if not all(part.isidentifier() for part in parts):
            raise ValueError(f'Invalid relocated package: {name}')
        package_source = source.joinpath(*parts)
        if not (package_source / '__init__.py').is_file():
            raise ValueError(f'Relocated package has no current source: {name}')
        destination = lib_dir / name
        if destination.is_symlink():
            destination.unlink()
        elif destination.exists():
            shutil.rmtree(destination)
        shutil.copytree(package_source, destination,
                        ignore=shutil.ignore_patterns('__pycache__', '*.pyc'))
        refreshed.append(name)
    return refreshed
