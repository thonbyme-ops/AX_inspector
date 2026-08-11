"""고정 계정 생성/조회용 CLI (이슈 #3 - 회원가입 UI는 없음, 관리자가 터미널에서 실행).

사용법:
  python manage_accounts.py add <username> <password> [--company "업체명"] [--display-name "이름"]
  python manage_accounts.py list

--company를 생략하면 관리자 계정(전체 업체 데이터 조회 가능)으로 생성된다.
"""
import argparse
import sqlite3
import sys

import db


def main():
    parser = argparse.ArgumentParser(description="AX_inspector 계정 관리")
    sub = parser.add_subparsers(dest="command", required=True)

    add_p = sub.add_parser("add", help="계정 생성")
    add_p.add_argument("username")
    add_p.add_argument("password")
    add_p.add_argument("--company", default=None, help="소속 업체명 (생략 시 관리자 계정)")
    add_p.add_argument("--display-name", default=None, help="표시 이름")

    sub.add_parser("list", help="계정 목록 조회")

    args = parser.parse_args()
    db.init_db()

    if args.command == "add":
        try:
            db.create_account(args.username, args.password, display_name=args.display_name, company=args.company)
        except sqlite3.IntegrityError:
            print(f"이미 존재하는 아이디입니다: {args.username}", file=sys.stderr)
            sys.exit(1)
        role = f"업체: {args.company}" if args.company else "관리자"
        print(f"계정 생성 완료: {args.username} ({role})")
    elif args.command == "list":
        accounts = db.list_accounts()
        if not accounts:
            print("생성된 계정이 없습니다.")
            return
        for acc in accounts:
            role = acc["company"] or "관리자"
            print(f"[{acc['id']}] {acc['username']} - {acc['display_name'] or ''} ({role})")


if __name__ == "__main__":
    main()
