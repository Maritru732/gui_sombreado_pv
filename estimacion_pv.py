import numpy as np

# Constantes
k = 1.3806503e-23
q = 1.60217646e-19


def estimar_parametros_pv(
    Voc, Isc, Vmp, Imp, Kv, Ki,
    Ns, T, a1, a2
):

    Voc =  Voc / Ns
    Vmp =  Vmp / Ns
    T = T + 273.15
    Vth = T * (k / q)
    p = a1 + a2
    Pmax_exp = Vmp * Imp

    Tsc = 25 + 273.15
    G = 1000
    Gsc = 1000
    deltaT = T - Tsc

    V_2 = np.linspace(0, Voc, 300)
    V_2 = np.sort(np.insert(V_2, 0, Vmp))

    Iph = (Isc + Ki * deltaT) * (G / Gsc)
    Io = (Isc + Ki * deltaT) / (
        np.exp((Voc + Kv * deltaT) / (((a1 + a2) / p) * Vth)) - 1
    )

    Rs = 0.0
    Rp = Vmp / (Isc - Vmp) - (Voc - Vmp) / Imp

    tol = 1e-3
    hist_Rs = []
    hist_Rp = []
    hist_error = []

    def corriente_pv(V, tol, max_iter=10):

        I_resultado = np.zeros(len(V))
        def f(I):
            term1 = np.exp((Vi + I * Rs) / Vth)
            term2 = np.exp((Vi + I * Rs) / ((p - 1) * Vth))
            return Iph - Io * (term1 + term2 + 2) - (Vi + I * Rs)/abs(Rp) - I

        def f_derivada(I):
            dterm1 = (Rs / Vth) * np.exp((Vi + I * Rs) / Vth)
            dterm2 = (Rs / ((p - 1) * Vth)) * np.exp((Vi + I * Rs) / ((p - 1) * Vth))
            return -Io * (dterm1 + dterm2) - (Rs / abs(Rp)) - 1

        for i, Vi in enumerate(V):
            # Valor inicial
            I = 0.

            # Método Newton-Raphson
            for _ in range(max_iter):
                f_val = f(I)
                f_der = f_derivada(I)
                if abs(f_der) < 1e-12:
                    raise ValueError(f"Derivada demasiado pequeña en V = {Vi}")
                I_new = I - f_val / f_der
                if abs(I_new - I) < tol:
                    break
                I = I_new
            I_resultado[i] = I

        return I_resultado

    error = 1.0
    cont_2 = 0
    while error > tol:
        cont_2 = cont_2 + 1
        I_2 = corriente_pv(V_2, tol)
        idx = np.argmax(V_2 * I_2)
        Pcalc = V_2[idx] * I_2[idx]
        error = abs(Pcalc - Pmax_exp)

        hist_Rs.append(Rs)
        hist_Rp.append(Rp)
        hist_error.append(error)

        Rs = round(Rs + 0.001, 6) #Redondear a 6 digitos
        Rp = (Vmp+Imp*Rs)/(Iph-Io*(np.exp((Vmp+Imp*Rs)/Vth)+np.exp((Vmp+Imp*Rs)/((p-1)*Vth))+2)-Pmax_exp/Vmp)
        if Rp<0 and cont_2>2:
            break

    idx_min = np.argmin(hist_error)
    Rs_final = hist_Rs[idx_min]
    Rp_final = hist_Rp[idx_min]

    I_final = corriente_pv(V_2,tol)
    P_final = V_2 * I_final

    idx_max = np.argmax(P_final)

    resultados = {
        "Corriente de saturación diodo 1 (Io1)": Io,
        "Corriente de saturación diodo 2 (Io2)": Io,
        "Fotocorriente (Iph)": Iph,
        "Resistencia serie (Rs)": Rs_final,
        "Resistencia paralelo (Rp)": Rp_final,
        "Potencia (Pmpp)": P_final[idx_max],
        "Voltaje (Vmpp)": V_2[idx_max],
        "Corriente (Impp)": I_final[idx_max],
    }

    datos_grafica = {
        "V_2": V_2,
        "I_final": I_final,
        "Vmp": Vmp,
        "Imp": Imp
    }

    return resultados, datos_grafica
