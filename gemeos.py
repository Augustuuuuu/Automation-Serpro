import socket
import re
from math import isqrt
import sys
import time

HOST = "desafio.ctfsecurityweek.serpro"
PORT = 10007
RECV_TIMEOUT = 10.0  # segundos para recv

# regexes para extrair N, e e c das linhas do servidor
re_N = re.compile(r'N:\s*([0-9]+)')
re_e = re.compile(r'e:\s*([0-9]+)')
re_c = re.compile(r'c(?:ipher|ifrada|ifrado)?[:)]*\s*([0-9]+)')  # cobre variações

def is_perfect_square(x: int) -> bool:
    if x < 0:
        return False
    r = isqrt(x)
    return r * r == x

def factor_twin_primes(N: int):
    # tenta Fermat simples: a = ceil(sqrt(N)), b^2 = a^2 - N
    a = isqrt(N)
    if a * a < N:
        a += 1

    b2 = a * a - N
    if not is_perfect_square(b2):
        # tenta alguns aumentos pequenos (normalmente não necessário para primos gêmeos)
        found = False
        for k in range(1, 5000):
            ak = a + k
            b2k = ak * ak - N
            if is_perfect_square(b2k):
                a = ak
                b2 = b2k
                found = True
                break
        if not found:
            raise ValueError("Não encontrei b^2 perfeito; N pode não ter primos tão próximos.")
    b = isqrt(b2)
    p = a - b
    q = a + b
    if p * q != N:
        raise ValueError("Fatoração inválida (p*q != N).")
    return p, q

def modinv(e, phi):
    # usa pow builtin (Python 3.8+)
    return pow(e, -1, phi)

def decrypt_flag_from_Nec(n_str, e_str, c_str):
    N = int(n_str)
    e = int(e_str)
    c = int(c_str)
    p, q = factor_twin_primes(N)
    phi = (p - 1) * (q - 1)
    d = modinv(e, phi)
    m = pow(c, d, N)
    blen = (m.bit_length() + 7) // 8
    if blen == 0:
        return b''
    return m.to_bytes(blen, 'big')

def readline_timeout(f):
    # f é o file-like object do socket.makefile('rwb' or 'r', encoding)
    try:
        line = f.readline()
        if line == '':
            # EOF
            return None
        return line
    except Exception:
        return None

def run():
    print(f"Conectando em {HOST}:{PORT} ...")
    s = socket.create_connection((HOST, PORT), timeout=RECV_TIMEOUT)
    # usar makefile para leitura por linha
    f = s.makefile('rw', encoding='utf-8', newline='\n')
    try:
        # ler cabeçalho inicial até prompt
        buffer = []
        while True:
            line = f.readline()
            if not line:
                print("Conexão fechada pelo servidor.")
                return
            print(line.rstrip())
            buffer.append(line)
            # esperar o prompt onde o servidor perguntou "digite 'pronto'..." ou similar
            if re.search(r'digite\s*"[Pp]ronto"', line) or re.search(r'digite\s+"pronto"', line):
                break
            # caso o prompt venha sem essa frase, também vamos romper se aparecer input prompt com '?'
            if line.strip().endswith(':') or line.strip().endswith('?'):
                break

        # enviar "pronto"
        print('Enviando: pronto')
        f.write('pronto\n')
        f.flush()

        # agora entra loop dos desafios; parsearemos N, e, c por bloco
        desafio_idx = 0
        while True:
            # ler até encontrarmos N, e, c
            N_val = None
            e_val = None
            c_val = None
            # ler linhas até acharmos c ou até timeout/EOF
            start_time = time.time()
            while True:
                line = f.readline()
                if not line:
                    print("Servidor fechou a conexão.")
                    return
                print(line.rstrip())

                # checar se encerramento ou "Parabéns"
                if "Parabéns" in line or "Parab" in line:
                    print("Recebido fim dos desafios.")
                    return
                # extrair N,e,c
                if N_val is None:
                    m = re_N.search(line)
                    if m:
                        N_val = m.group(1)
                if e_val is None:
                    m = re_e.search(line)
                    if m:
                        e_val = m.group(1)
                if c_val is None:
                    m = re_c.search(line)
                    if m:
                        c_val = m.group(1)

                # Se encontramos N,e,c, fazemos a descriptografia e enviamos a flag
                if N_val and e_val and c_val:
                    desafio_idx += 1
                    print(f"Desafio {desafio_idx}: extraídos N, e e c. Calculando flag...")
                    try:
                        msg_bytes = decrypt_flag_from_Nec(N_val, e_val, c_val)
                        # tentar decodificar, mas enviar como string (utf-8)
                        try:
                            flag_str = msg_bytes.decode('utf-8')
                        except Exception:
                            # se não decodificar, enviar repr hex ou bytes direto - porém servidor espera texto, então mandar latin-1 fallback
                            flag_str = msg_bytes.decode('latin-1')
                        print("Flag calculada:", flag_str)
                        # enviar resposta (com newline)
                        print("Enviando flag para o servidor...")
                        f.write(flag_str + '\n')
                        f.flush()
                    except Exception as ex:
                        print("Erro ao calcular flag:", ex)
                        print("Enviando resposta vazia para abortar.")
                        f.write('\n')
                        f.flush()
                    # após enviar, continuar lendo próximo bloco
                    break

                # safety timeout para não travar
                if time.time() - start_time > 30:
                    print("Timeout lendo do servidor (30s).")
                    return

    finally:
        try:
            f.close()
        except Exception:
            pass
        try:
            s.close()
        except Exception:
            pass

if __name__ == "__main__":
    try:
        run()
    except KeyboardInterrupt:
        print("\nInterrompido pelo usuário.")
        sys.exit(0)
