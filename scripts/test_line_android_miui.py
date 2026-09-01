from __future__ import annotations

import io
import sqlite3
import tarfile
import tempfile
import unittest
from argparse import Namespace
from pathlib import Path

import extract_line_android_miui
import reconcile


class LineAndroidMiuiTests(unittest.TestCase):
    def _build_backup(self, root: Path) -> tuple[Path, bytes]:
        source = root / "miui"
        source.mkdir()
        line_db = root / "naver_line"
        connection = sqlite3.connect(line_db)
        connection.executescript(
            """
            CREATE TABLE groups (id TEXT PRIMARY KEY, name TEXT);
            CREATE TABLE chat_history (
                id INTEGER PRIMARY KEY,
                server_id TEXT,
                type INTEGER,
                chat_id TEXT,
                from_mid TEXT,
                content TEXT,
                created_time TEXT,
                status INTEGER,
                attachement_type INTEGER,
                parameter TEXT
            );
            INSERT INTO groups VALUES ('c-small', '测试小额群');
            INSERT INTO groups VALUES ('c-other', '普通群');
            INSERT INTO chat_history VALUES
                (1, '100', 1, 'c-small', 'u-staff', '报价 32.5', '1780000000000', 3, 0, ''),
                (2, '101', 1, 'c-small', 'u-customer', '', '1780000060000', 3, 1,
                 'message_relation_server_message_id\t100'),
                (3, '102', 1, 'c-small', 'u-staff', '', '1780000120000', 3, 1, ''),
                (4, '103', 13, 'c-small', NULL, NULL, '1780000180000', 3, 18,
                 'LOC_KEY\tC_MI\tLOC_ARGS\tu-customer');
            """
        )
        connection.commit()
        connection.close()

        contact_db = root / "contact"
        connection = sqlite3.connect(contact_db)
        connection.executescript(
            """
            CREATE TABLE contacts (mid TEXT PRIMARY KEY, profile_name TEXT, overridden_name TEXT);
            INSERT INTO contacts VALUES ('u-staff', 'QQ～财务3（休息）', NULL);
            INSERT INTO contacts VALUES ('u-customer', 'Alice', NULL);
            """
        )
        connection.commit()
        connection.close()

        jpeg = b"\xff\xd8\xff\xe0fixture-jpeg\xff\xd9"
        backup = source / "LINE(jp.naver.line.android).bak"
        header = (
            b"MIUI BACKUP\n2\n"
            b"jp.naver.line.android LINE\n102\n0\n"
            b"ANDROID BACKUP\n5\n0\nnone\n"
        )
        with backup.open("wb") as handle:
            handle.write(header)
            with tarfile.open(fileobj=handle, mode="w") as archive:
                archive.add(
                    line_db,
                    arcname="apps/jp.naver.line.android/db/naver_line",
                )
                archive.add(
                    contact_db,
                    arcname="apps/jp.naver.line.android/db/contact",
                )
                info = tarfile.TarInfo(
                    "apps/jp.naver.line.android/ef/chats/c-small/messages/2"
                )
                info.size = len(jpeg)
                archive.addfile(info, io.BytesIO(jpeg))
        (source / "descript.xml").write_text(
            "<MIUI-backup><packages><package><packageName>jp.naver.line.android</packageName>"
            "</package></packages></MIUI-backup>",
            encoding="utf-8",
        )
        return backup, jpeg

    def test_start_discovers_and_normalizes_miui_backup(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            backup, jpeg = self._build_backup(root)
            groups = extract_line_android_miui.available_groups_for_backup(backup.parent)
            self.assertEqual([group["group_name"] for group in groups], ["普通群", "测试小额群"])

            work = root / "run"
            report = reconcile.start_run(
                Namespace(
                    inputs=[backup.parent],
                    work=work,
                    contains="小额",
                    timezone="Asia/Bangkok",
                    roster=Path(__file__).resolve().parent.parent / "config" / "roster.yaml",
                    line_backup=[],
                    line_android_backup=[],
                    line_self_name="LINE_SELF",
                )
            )
            run = reconcile._load_run(work)
            normalized = reconcile._load_snapshot(work, run)
            self.assertEqual(report["groups"], 1)
            self.assertEqual(run["line_android_backups"], [str(backup.resolve())])
            group = normalized["groups"][0]
            self.assertEqual(group["group_key"], "line:c-small")
            self.assertEqual(group["platform"], "LINE")
            self.assertEqual(len(group["messages"]), 4)
            self.assertEqual(group["messages"][0]["role"], "内部人员")
            self.assertEqual(
                group["messages"][1]["reply_to_message_id"],
                group["messages"][0]["message_id"],
            )
            available = group["messages"][1]["media"][0]
            self.assertEqual(available["availability"], "available")
            self.assertEqual(Path(available["path"]).read_bytes(), jpeg)
            missing = group["messages"][2]["media"][0]
            self.assertEqual(missing["availability"], "missing")
            self.assertTrue(group["messages"][3]["excluded_from_accounting"])

    def test_compressed_backup_is_rejected_clearly(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / "LINE(jp.naver.line.android).bak"
            path.write_bytes(
                b"MIUI BACKUP\n2\njp.naver.line.android LINE\n102\n0\n"
                b"ANDROID BACKUP\n5\n1\nnone\n"
            )
            with self.assertRaisesRegex(ValueError, "compression flag 0"):
                extract_line_android_miui.backup_header(path)


if __name__ == "__main__":
    unittest.main()
