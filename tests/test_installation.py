import json
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'scripts'))
import install_skill as install

class InstallTests(unittest.TestCase):
    def fixture(self,root):
        home=root/'user';ch=home/'.codex'
        old=home/'.agents/skills/documents';old.mkdir(parents=True)
        (old/'SKILL.md').write_text('---\nname: documents\n---\nOld user skill')
        plugin=ch/'plugins/cache/vendor/plugin/1.0/skills/documents';plugin.mkdir(parents=True)
        (plugin/'SKILL.md').write_text('---\nname: documents\n---\nOriginal plugin skill')
        another=ch/'plugins/cache/vendor/plugin/1.0/skills/other';another.mkdir()
        (another/'SKILL.md').write_text('---\nname: other\n---\nOther skill')
        config='model = "example"\n\n[[skills.config]]\npath = '+json.dumps(str(plugin/'SKILL.md'))+'\nenabled = true # keep comment\n\n[unrelated]\nvalue = 42\n'
        (ch/'config.toml').write_text(config)
        return home,ch,old,plugin,another,config

    def test_install_preserves_plugin_and_updates_only_matching_config(self):
        with tempfile.TemporaryDirectory() as td:
            home,ch,old,plugin,other,original=self.fixture(Path(td))
            result=install.install(ROOT,home,ch,skip_deps=True)
            self.assertEqual((plugin/'SKILL.md').read_text(),'---\nname: documents\n---\nOriginal plugin skill')
            self.assertTrue((other/'SKILL.md').exists())
            self.assertTrue((old/'scripts/user_formatting.py').exists())
            backup=Path(result['backups'][0]['backup'])
            self.assertIn('Old user skill',(backup/'SKILL.md').read_text())
            self.assertEqual((Path(result['backup_directory'])/'config.toml').read_text(),original)
            text=(ch/'config.toml').read_text()
            self.assertIn('enabled = false # keep comment',text)
            self.assertIn('[unrelated]\nvalue = 42',text)
            self.assertTrue((old/'install-state.json').exists())
            self.assertFalse((old/'.git').exists())
            again=install.install(old,home,ch,skip_deps=True)
            self.assertNotEqual(result['backup_directory'],again['backup_directory'])
            self.assertEqual((ch/'config.toml').read_text().count('[[skills.config]]'),2)

    def test_dry_run_has_no_side_effects(self):
        with tempfile.TemporaryDirectory() as td:
            home,ch,old,plugin,other,original=self.fixture(Path(td))
            result=install.install(ROOT,home,ch,skip_deps=True,dry_run=True)
            self.assertTrue(result['dry_run'])
            self.assertIn('Old user skill',(old/'SKILL.md').read_text())
            self.assertEqual((ch/'config.toml').read_text(),original)
            self.assertFalse((ch/'skill-backups').exists())

    def test_rollback_restores_original_after_copy_failure(self):
        with tempfile.TemporaryDirectory() as td:
            home,ch,old,plugin,other,original=self.fixture(Path(td))
            real=install.shutil.copytree
            def fail(source,dest,*args,**kwargs):
                if install.canonical(dest)==install.canonical(old):
                    old.mkdir();(old/'partial').write_text('partial')
                    raise OSError('test copy failure')
                return real(source,dest,*args,**kwargs)
            with patch.object(install.shutil,'copytree',side_effect=fail):
                with self.assertRaises(OSError):install.install(ROOT,home,ch,skip_deps=True)
            self.assertIn('Old user skill',(old/'SKILL.md').read_text())
            self.assertFalse((old/'partial').exists())
            self.assertEqual((ch/'config.toml').read_text(),original)

    def test_existing_directory_path_and_unrelated_entries(self):
        with tempfile.TemporaryDirectory() as td:
            root=Path(td);p=root/'plugin/documents/SKILL.md'
            text='[[skills.config]]\npath = '+json.dumps(str(p.parent))+'\nenabled = true\n\n[[skills.config]]\npath = '+json.dumps(str(root/'other/SKILL.md'))+'\nenabled = true\n'
            changed=install.update_config(text,{p:False})
            self.assertEqual(changed.count('enabled = false'),1)
            self.assertEqual(changed.count('enabled = true'),1)

if __name__=='__main__':unittest.main()
