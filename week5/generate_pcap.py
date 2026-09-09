import struct
import time
import os

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_PCAP = os.path.join(SCRIPT_DIR, "traffic_analysis_0x02.pcap")

# Helper for IPv4 checksum
def ip_checksum(header: bytes) -> int:
    if len(header) % 2 == 1:
        header += b'\x00'
    s = sum(struct.unpack('!%dH' % (len(header) // 2), header))
    s = (s >> 16) + (s & 0xffff)
    s += s >> 16
    return ~s & 0xffff

def mac_to_bytes(mac: str) -> bytes:
    return bytes.fromhex(mac.replace(':', ''))

def ip_to_bytes(ip: str) -> bytes:
    return bytes(map(int, ip.split('.')))

class PcapBuilder:
    def __init__(self):
        self.packets = []
        self.base_time = 1757433600.0  # Simulated timestamp

    def add_raw_packet(self, data: bytes, delta_sec: float):
        t = self.base_time + delta_sec
        sec = int(t)
        usec = int((t - sec) * 1_000_000)
        hdr = struct.pack('<IIII', sec, usec, len(data), len(data))
        self.packets.append(hdr + data)

    def add_arp(self, delta_sec: float, op: int, src_mac: str, src_ip: str, dst_mac: str, dst_ip: str):
        # Ethernet Header
        eth_dst = mac_to_bytes(dst_mac if op == 2 else "ff:ff:ff:ff:ff:ff")
        eth_src = mac_to_bytes(src_mac)
        eth = eth_dst + eth_src + struct.pack('!H', 0x0806)
        
        # ARP Payload
        arp = struct.pack('!HHBBH', 1, 0x0800, 6, 4, op)
        arp += mac_to_bytes(src_mac) + ip_to_bytes(src_ip)
        arp += mac_to_bytes(dst_mac if op == 2 else "00:00:00:00:00:00") + ip_to_bytes(dst_ip)
        self.add_raw_packet(eth + arp, delta_sec)

    def add_dns(self, delta_sec: float, src_ip: str, dst_ip: str, src_port: int, domain: str, is_response=False, resolved_ip="10.0.0.88"):
        eth = mac_to_bytes("00:0c:29:ab:cd:ef") + mac_to_bytes("00:1a:2b:3c:4d:5e") + struct.pack('!H', 0x0800)
        
        # DNS Query Payload
        tx_id = 0x1337
        flags = 0x8180 if is_response else 0x0100
        dns_hdr = struct.pack('!HHHHHH', tx_id, flags, 1, 1 if is_response else 0, 0, 0)
        
        qname = b''.join(bytes([len(part)]) + part.encode('ascii') for part in domain.split('.')) + b'\x00'
        dns_payload = dns_hdr + qname + struct.pack('!HH', 1, 1) # Type A, Class IN
        if is_response:
            ans = struct.pack('!HHHLH', 0xc00c, 1, 1, 300, 4) + ip_to_bytes(resolved_ip)
            dns_payload += ans

        udp_len = 8 + len(dns_payload)
        udp_hdr = struct.pack('!HHHH', src_port, 53 if not is_response else src_port, udp_len, 0)
        
        ip_total_len = 20 + udp_len
        ip_hdr = struct.pack('!BBHHHBBH', 0x45, 0, ip_total_len, 0x1234, 0x4000, 64, 17, 0)
        ip_hdr += ip_to_bytes(src_ip) + ip_to_bytes(dst_ip)
        chk = ip_checksum(ip_hdr)
        ip_hdr = ip_hdr[:10] + struct.pack('!H', chk) + ip_hdr[12:]

        self.add_raw_packet(eth + ip_hdr + udp_hdr + dns_payload, delta_sec)

    def add_tcp(self, delta_sec: float, src_ip: str, dst_ip: str, src_port: int, dst_port: int,
                seq: int, ack: int, flags: int, payload: bytes = b''):
        eth = mac_to_bytes("00:0c:29:ab:cd:ef") + mac_to_bytes("00:1a:2b:3c:4d:5e") + struct.pack('!H', 0x0800)
        
        tcp_hdr_len = 20
        tcp_hdr = struct.pack('!HHIIBBHHH', src_port, dst_port, seq, ack, (tcp_hdr_len // 4) << 4, flags, 64240, 0, 0)
        
        ip_total_len = 20 + tcp_hdr_len + len(payload)
        ip_hdr = struct.pack('!BBHHHBBH', 0x45, 0, ip_total_len, (seq & 0xffff), 0x4000, 64, 6, 0)
        ip_hdr += ip_to_bytes(src_ip) + ip_to_bytes(dst_ip)
        chk = ip_checksum(ip_hdr)
        ip_hdr = ip_hdr[:10] + struct.pack('!H', chk) + ip_hdr[12:]
        
        self.add_raw_packet(eth + ip_hdr + tcp_hdr + payload, delta_sec)

    def build(self) -> bytes:
        # Standard libpcap global header
        global_hdr = struct.pack('<IHHiIII', 0xa1b2c3d4, 2, 4, 0, 0, 65535, 1)
        return global_hdr + b''.join(self.packets)

def generate():
    builder = PcapBuilder()
    
    # 1. ARP Request & Reply (Gateway discovery)
    builder.add_arp(0.01, 1, "00:1a:2b:3c:4d:5e", "192.168.1.42", "00:00:00:00:00:00", "192.168.1.1")
    builder.add_arp(0.02, 2, "00:0c:29:ab:cd:ef", "192.168.1.1", "00:1a:2b:3c:4d:5e", "192.168.1.42")
    
    # 2. DNS Resolution for internal gateway
    builder.add_dns(0.15, "192.168.1.42", "192.168.1.1", 54312, "gateway.baysec.sfbu", False)
    builder.add_dns(0.18, "192.168.1.1", "192.168.1.42", 54312, "gateway.baysec.sfbu", True, "10.0.0.88")
    
    # 3. Initial TCP 3-Way Handshake to Web Server (Port 80)
    builder.add_tcp(0.30, "192.168.1.42", "10.0.0.88", 49820, 80, 1000, 0, 0x02) # SYN
    builder.add_tcp(0.33, "10.0.0.88", "192.168.1.42", 80, 49820, 5000, 1001, 0x12) # SYN-ACK
    builder.add_tcp(0.34, "192.168.1.42", "10.0.0.88", 49820, 80, 1001, 5001, 0x10) # ACK
    
    # 4. HTTP GET /healthcheck
    http_get = (
        b"GET /healthcheck HTTP/1.1\r\n"
        b"Host: gateway.baysec.sfbu\r\n"
        b"User-Agent: Mozilla/5.0 (X11; Linux x86_64; rv:128.0) Gecko/20100101 Firefox/128.0\r\n"
        b"Accept: */*\r\n\r\n"
    )
    builder.add_tcp(0.40, "192.168.1.42", "10.0.0.88", 49820, 80, 1001, 5001, 0x18, http_get)
    
    http_resp1 = (
        b"HTTP/1.1 200 OK\r\n"
        b"Server: nginx/1.24.0 (Ubuntu)\r\n"
        b"Content-Type: text/plain\r\n"
        b"Content-Length: 18\r\n\r\n"
        b"SYSTEM ONLINE: OK\n"
    )
    builder.add_tcp(0.45, "10.0.0.88", "192.168.1.42", 80, 49820, 5001, 1001 + len(http_get), 0x18, http_resp1)
    
    # 5. Background encrypted TLS connection (distractor)
    builder.add_tcp(0.60, "192.168.1.42", "104.21.48.22", 51220, 443, 2000, 0, 0x02) # SYN
    builder.add_tcp(0.64, "104.21.48.22", "192.168.1.42", 443, 51220, 8000, 2001, 0x12) # SYN-ACK
    builder.add_tcp(0.65, "192.168.1.42", "104.21.48.22", 51220, 443, 2001, 8001, 0x10) # ACK
    tls_client_hello = b"\x16\x03\x01\x00\x98\x01\x00\x00\x94\x03\x03" + os.urandom(32) + b"\x00\x00\x20" + os.urandom(32)
    builder.add_tcp(0.70, "192.168.1.42", "104.21.48.22", 51220, 443, 2001, 8001, 0x18, tls_client_hello)

    # 6. Secondary TCP Handshake to Port 8080 (Telemetry & Operator Uplink)
    builder.add_tcp(1.20, "192.168.1.42", "10.0.0.88", 55432, 8080, 3000, 0, 0x02) # SYN
    builder.add_tcp(1.23, "10.0.0.88", "192.168.1.42", 8080, 55432, 9000, 3001, 0x12) # SYN-ACK
    builder.add_tcp(1.24, "192.168.1.42", "10.0.0.88", 55432, 8080, 3001, 9001, 0x10) # ACK

    # 7. CRITICAL PACKET: HTTP POST /api/v2/telemetry/uplink WITH THE FLAG!
    flag_body = (
        b"station_id=SFBU-LAB-01&"
        b"operator_callsign=0xGhost&"
        b"auth_token=sfbu{w1r3sh4rk_p4ck3t_sn1ff3r_2026}&"
        b"status=beacon_active"
    )
    http_post = (
        b"POST /api/v2/telemetry/uplink HTTP/1.1\r\n"
        b"Host: gateway.baysec.sfbu:8080\r\n"
        b"User-Agent: BaySec-Telemetry-Uplink/2.4 (x86_64; Linux)\r\n"
        b"Content-Type: application/x-www-form-urlencoded\r\n"
        b"Authorization: Basic c2ZidV9hZG1pbjpoNGNrM3JfcDRzc3cwcmQ=\r\n"
        b"Content-Length: " + str(len(flag_body)).encode('ascii') + b"\r\n\r\n" + flag_body
    )
    builder.add_tcp(1.30, "192.168.1.42", "10.0.0.88", 55432, 8080, 3001, 9001, 0x18, http_post)
    
    # 8. Server Acknowledgment & 200 OK Response
    http_resp2 = (
        b"HTTP/1.1 200 OK\r\n"
        b"Server: BaySec-Gatekeeper/1.0\r\n"
        b"Content-Type: application/json\r\n"
        b"Content-Length: 64\r\n\r\n"
        b"{\"status\":\"granted\",\"msg\":\"Operator beacon authenticated. Welcome.\"}\n"
    )
    builder.add_tcp(1.35, "10.0.0.88", "192.168.1.42", 8080, 55432, 9001, 3001 + len(http_post), 0x18, http_resp2)
    
    # 9. TCP Teardown (FIN/ACK handshakes)
    builder.add_tcp(1.80, "192.168.1.42", "10.0.0.88", 55432, 8080, 3001 + len(http_post), 9001 + len(http_resp2), 0x11)
    builder.add_tcp(1.82, "10.0.0.88", "192.168.1.42", 8080, 55432, 9001 + len(http_resp2), 3002 + len(http_post), 0x10)
    builder.add_tcp(1.83, "10.0.0.88", "192.168.1.42", 8080, 55432, 9001 + len(http_resp2), 3002 + len(http_post), 0x11)
    builder.add_tcp(1.85, "192.168.1.42", "10.0.0.88", 55432, 8080, 3002 + len(http_post), 9002 + len(http_resp2), 0x10)

    pcap_data = builder.build()
    with open(OUTPUT_PCAP, "wb") as f:
        f.write(pcap_data)
    print(f"Generated authentic PCAP file with {len(builder.packets)} packets at: {OUTPUT_PCAP}")

if __name__ == "__main__":
    generate()
