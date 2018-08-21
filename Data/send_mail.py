import logging
import smtplib


class Mail:
    ip = 587
    server_status = False

    def __init__(self, user_id, user_pass):
        if self.start_service():
            try:
                server = smtplib.SMTP('smtp.gmail.com', self.ip)
                logging.info("Starting SMTP libraries as server {} using ip {}".format(server, self.ip))
                server.starttls()
                server.login(user_id, user_pass)
                logging.info("Started server using address {}".format(user_id))
                self.server_status = True
                self.server = server
                self.user_id = user_id

            except Exception as exp:
                logging.error("____INVALID-ID/PASS____")

        else:
            logging.warn("Unable to start services...")

    @staticmethod
    def start_service():
        try:
            ip = 587
            server = smtplib.SMTP('smtp.gmail.com', ip)
            logging.debug("Testing SMTP libraries as server {} using ip {}".format(server, ip))
            server.starttls()
            return True

        except Exception as exp:
            logging.error("While starting server an error was raised => {}".format(exp))
            return False


    def sendmail(self, address, msg, sbj = "NYMUN Confirmation Mail"):
        other_address = self.user_id
        try:
            try:
                message = 'Subject: {}\n\n{}'.format(sbj, msg)
                self.server.sendmail(self.user_id, address, message)
                logging.info("Sent mail at address {}".format(address))
                return True

            except Exception as exc:
                logging.error("While sending mail an error was raised => {}".format(exc))
                message = 'Subject: {}\n\n{}'.format(sbj, msg)
                self.server.sendmail(self.user_id, other_address, message)
                logging.info("Sent mail at address {}".format(other_address))
                return "____INVALID-MAIL-ADDRESS____"

        except Exception as exp:
            logging.error("While sending mail error was raised again => {}".format(exp))
            self.server_status = False
            self.server.quit()
            logging.info("Server closed")
            return "____CONNECTION-TIMEDOUT_____"
