import smtplib
import socket

def test_smtp_connection(host, port, use_ssl=False):
    print(f"Testing {host}:{port} (SSL: {use_ssl})...")
    try:
        if use_ssl:
            server = smtplib.SMTP_SSL(host, port, timeout=5)
        else:
            server = smtplib.SMTP(host, port, timeout=5)
        print(f"✅ Successfully connected to {host}:{port}")
        server.quit()
        return True
    except Exception as e:
        print(f"❌ Failed to connect to {host}:{port}: {e}")
        return False

if __name__ == "__main__":
    test_smtp_connection('smtp.gmail.com', 465, use_ssl=True)
    test_smtp_connection('smtp.gmail.com', 587, use_ssl=False)
