import argparse
import json
from functools import partial
from http import HTTPStatus
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from excel_io import data_from_workbook, workbook_bytes


STATIC_DIR = Path(__file__).parent / "static"
MAX_BODY_SIZE = 20 * 1024 * 1024


class AppHandler(SimpleHTTPRequestHandler):
	def _send_json(self, status, payload):
		body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
		self.send_response(status)
		self.send_header("Content-Type", "application/json; charset=utf-8")
		self.send_header("Content-Length", str(len(body)))
		self.send_header("Cache-Control", "no-store")
		self.end_headers()
		self.wfile.write(body)

	def _read_body(self):
		try:
			length = int(self.headers.get("Content-Length", "0"))
		except ValueError as error:
			raise ValueError("Некорректный размер запроса") from error
		if length < 0 or length > MAX_BODY_SIZE:
			raise ValueError("Размер файла превышает 20 МБ")
		body = self.rfile.read(length)
		if len(body) != length:
			raise ValueError("Запрос передан не полностью")
		return body

	def do_POST(self):
		if self.path not in {"/api/import.xlsx", "/api/export.xlsx"}:
			self.send_error(HTTPStatus.NOT_FOUND)
			return

		try:
			body = self._read_body()
			if self.path == "/api/import.xlsx":
				self._send_json(HTTPStatus.OK, data_from_workbook(body))
				return

			data = json.loads(body)
			if not isinstance(data, dict):
				raise ValueError("Ожидается объект JSON")
			workbook = workbook_bytes(data)
			self.send_response(HTTPStatus.OK)
			self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
			self.send_header("Content-Disposition", 'attachment; filename="experimental-data.xlsx"')
			self.send_header("Content-Length", str(len(workbook)))
			self.end_headers()
			self.wfile.write(workbook)
		except (ValueError, json.JSONDecodeError) as error:
			self._send_json(HTTPStatus.BAD_REQUEST, {"error": str(error)})
		except Exception as error:
			self.log_error("Excel API error: %s", error)
			self._send_json(HTTPStatus.INTERNAL_SERVER_ERROR, {"error": "Не удалось обработать Excel-файл"})


def main():
	parser = argparse.ArgumentParser(description="Experimental paint product data entry app")
	parser.add_argument("--host", default="127.0.0.1")
	parser.add_argument("--port", type=int, default=8000)
	args = parser.parse_args()

	handler = partial(AppHandler, directory=str(STATIC_DIR))
	server = ThreadingHTTPServer((args.host, args.port), handler)
	print(f"Serving the app at http://{args.host}:{args.port}")
	try:
		server.serve_forever()
	except KeyboardInterrupt:
		print("\nStopping the app")
	finally:
		server.server_close()


if __name__ == "__main__":
	main()
