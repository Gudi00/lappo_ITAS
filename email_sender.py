"""
Модуль отправки электронной почты через Gmail SMTP.

Обеспечивает:
- Отправку писем через Gmail с App Password
- Прикрепление файлов (.docx и другие)
- Массовую рассылку
- Логирование отправки в базу данных
- Режим тестовой отправки (dry run)
"""

import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from typing import Optional

import config
from database import Database


class EmailSender:
    """Класс для отправки писем через Gmail SMTP."""

    def __init__(
        self,
        gmail_address: str = "",
        app_password: str = "",
        db: Optional[Database] = None,
        dry_run: bool = False
    ):
        """
        Аргументы:
            gmail_address: Адрес Gmail отправителя
            app_password: Пароль приложения Gmail
            db: Экземпляр Database для логирования
            dry_run: Если True — письма не отправляются, только логируются
        """
        self.gmail_address = gmail_address or config.GMAIL_ADDRESS
        self.app_password = app_password or config.GMAIL_APP_PASSWORD
        self.db = db
        self.dry_run = dry_run

    def _validate_config(self):
        """Проверяет наличие необходимых настроек."""
        if not self.gmail_address:
            raise ValueError(
                "Не указан адрес Gmail.\n"
                "Установите GMAIL_ADDRESS в файле .env"
            )
        if not self.app_password:
            raise ValueError(
                "Не указан пароль приложения Gmail.\n"
                "Установите GMAIL_APP_PASSWORD в файле .env\n"
                "Инструкция: https://support.google.com/accounts/answer/185833"
            )

    def send_email(
        self,
        to: str,
        subject: str,
        body: str,
        html_body: Optional[str] = None,
        attachment_path: Optional[Path] = None,
    ) -> bool:
        """
        Отправляет одно письмо.

        Аргументы:
            to: Адрес получателя
            subject: Тема письма
            body: Текст письма (plain text)
            html_body: HTML-версия текста (опционально)
            attachment_path: Путь к прикрепляемому файлу

        Возвращает:
            True если письмо отправлено успешно
        """
        self._validate_config()

        # Создаём запись в логе
        log_id = None
        if self.db:
            log_id = self.db.log_email(
                recipient=to,
                subject=subject,
                body_preview=body[:200] if body else "",
                status="dry_run" if self.dry_run else "sending",
                attachment_path=str(attachment_path) if attachment_path else ""
            )

        if self.dry_run:
            print(f"[ТЕСТ] Письмо для {to}: {subject}")
            if self.db and log_id:
                self.db.update_email_status(log_id, "dry_run")
            return True

        try:
            # Формируем письмо
            msg = MIMEMultipart()
            msg["From"] = self.gmail_address
            msg["To"] = to
            msg["Subject"] = subject

            # Текст письма
            if html_body:
                msg.attach(MIMEText(body, "plain", "utf-8"))
                msg.attach(MIMEText(html_body, "html", "utf-8"))
            else:
                msg.attach(MIMEText(body, "plain", "utf-8"))

            # Прикрепляем файл
            if attachment_path and Path(attachment_path).exists():
                file_path = Path(attachment_path)
                with open(file_path, "rb") as f:
                    attachment = MIMEApplication(f.read())
                    attachment.add_header(
                        "Content-Disposition", "attachment",
                        filename=file_path.name
                    )
                    msg.attach(attachment)

            # Отправляем
            with smtplib.SMTP(config.SMTP_SERVER, config.SMTP_PORT) as server:
                server.ehlo()
                server.starttls()
                server.ehlo()
                server.login(self.gmail_address, self.app_password)
                server.send_message(msg)

            # Обновляем лог
            if self.db and log_id:
                self.db.update_email_status(log_id, "sent")

            return True

        except smtplib.SMTPAuthenticationError:
            error_msg = (
                "Ошибка аутентификации Gmail.\n"
                "Проверьте:\n"
                "1. Правильность адреса Gmail\n"
                "2. Правильность пароля приложения\n"
                "3. Включена ли двухфакторная аутентификация"
            )
            if self.db and log_id:
                self.db.update_email_status(log_id, "error", error_msg)
            raise ConnectionError(error_msg)

        except smtplib.SMTPException as e:
            error_msg = f"Ошибка SMTP: {str(e)}"
            if self.db and log_id:
                self.db.update_email_status(log_id, "error", error_msg)
            raise ConnectionError(error_msg)

        except Exception as e:
            error_msg = f"Неожиданная ошибка при отправке: {str(e)}"
            if self.db and log_id:
                self.db.update_email_status(log_id, "error", error_msg)
            raise

    def send_bulk(
        self,
        recipients: list[dict],
        subject: str,
        body: str,
        html_body: Optional[str] = None,
    ) -> dict:
        """
        Массовая рассылка писем.

        Аргументы:
            recipients: Список словарей с ключами:
                - 'email': адрес получателя
                - 'attachment' (опционально): путь к файлу
            subject: Тема письма (общая для всех)
            body: Текст письма
            html_body: HTML-версия текста

        Возвращает:
            Словарь со статистикой: {'sent': N, 'failed': N, 'errors': [...]}
        """
        results = {"sent": 0, "failed": 0, "errors": []}

        for recipient in recipients:
            email = recipient.get("email", "")
            attachment = recipient.get("attachment")

            if not email:
                results["failed"] += 1
                results["errors"].append("Пустой адрес получателя")
                continue

            try:
                attachment_path = Path(attachment) if attachment else None
                self.send_email(
                    to=email,
                    subject=subject,
                    body=body,
                    html_body=html_body,
                    attachment_path=attachment_path
                )
                results["sent"] += 1
            except Exception as e:
                results["failed"] += 1
                results["errors"].append(f"{email}: {str(e)}")

        return results
