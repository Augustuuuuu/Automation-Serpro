
import socket
import re
import time

HOST = "desafio.ctfsecurityweek.serpro"
PORT = 10006
NTIMES_REQUIRED = 25

# ----------------- funções de untemper -----------------
def u32(x): return x & 0xFFFFFFFF

def undo_right_shift_xor(y, shift):
    res = 0
    for i in range(32):
        part = (y ^ (res >> shift)) & (1 << i)
        res |= part
    return u32(res)

def undo_left_shift_xor_mask(y, shift, mask):
    res = 0
    for i in range(32):
        bit = (y ^ ((res << shift) & mask)) & (1 << i)
        res |= bit
    return u32(res)

def untemper(y):
    y = undo_right_shift_xor(y, 18)
    y = undo_left_shift_xor_mask(y, 15, 0xEFC60000)
    y = undo_left_shift_xor_mask(y, 7, 0x9D2C5680)
    y = undo_right_shift_xor(y, 11)
    return u32(y)

# ----------------- implementação MT19937 -----------------
class MT19937:
    def __init__(self, state):
        self.mt = [u32(x) for x in state]
        self.index = 624

    def twist(self):
        for i in range(624):
            y = (self.mt[i] & 0x80000000) + (self.mt[(i+1)%624] & 0x7fffffff)
            self.mt[i] = self.mt[(i+397)%624] ^ (y >> 1) ^ (0x9908b0df if (y & 1) else 0)
        self.index = 0

    def extract_number(self):
        if self.index >= 624:
            self.twist()
        y = self.mt[self.index]
        y ^= (y >> 11)
        y ^= (y << 7) & 0x9D2C5680
        y ^= (y << 15) & 0xEFC60000
        y ^= (y >> 18)
        self.index += 1
        return u32(y)

# ----------------- coleta e envio de palpites -----------------
RE_NUM = re.compile(r"eu estava pensando no número:\s*([0-9]+)", re.IGNORECASE)

def recv_all(sock, timeout=0.5):
    sock.settimeout(timeout)
    data = b""
    try:
        while True:
            chunk = sock.recv(4096)
            if not chunk:
                break
            data += chunk
    except socket.timeout:
        pass
    return data.decode('utf-8', errors='ignore')

def collect_numbers(sock, n_needed):
    nums = []
    while len(nums) < n_needed:
        sock.sendall(b"0\n")  # palpite errado
        time.sleep(0.05)
        text = recv_all(sock, timeout=0.5)
        m = RE_NUM.search(text)
        if m:
            nums.append(int(m.group(1)))
            print(f"[coletado] {len(nums)}: {nums[-1]}")
    return nums

def main():
    print(f"Conectando a {HOST}:{PORT} ...")
    s = socket.create_connection((HOST, PORT))
    time.sleep(0.1)
    welcome = recv_all(s, timeout=1)
    print(welcome)

    # envia "pronto" para iniciar
    s.sendall(b"pronto\n")
    time.sleep(0.1)
    recv_all(s, timeout=0.5)

    # coleta 624 números para reconstruir MT19937
    print("Coletando 624 números do servidor...")
    nums = collect_numbers(s, 624)

    # untemper e criar MT local
    state = [untemper(n) for n in nums]
    mt = MT19937(state)

    # envia os 25 palpites corretos
    print("Enviando 25 palpites corretos...")
    for i in range(NTIMES_REQUIRED):
        nxt = mt.extract_number()
        s.sendall(f"{nxt}\n".encode())
        time.sleep(0.05)
        resp = recv_all(s, timeout=0.5)
        print(resp)

    # receber resposta final (flag)
    time.sleep(0.2)
    final = recv_all(s, timeout=1)
    print("=== FLAG ===")
    print(final)
    s.close()

if __name__ == "__main__":
    main()