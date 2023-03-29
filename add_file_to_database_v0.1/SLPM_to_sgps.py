def SLPM_to_sgps(SLPM):
    temperature_std=float(298.15) # standard temperature [K]
    P_std=float(101325.01) # standard pressure [Pa]
    mass_mol=float(0.02897) # molar mass mol of air [kg/mol]
    R=float(8.3144326) # molar gas constant [kgm2/s2Kmol]
    rho_std=(mass_mol*P_std)/(R*temperature_std) # standard density of air [kg/m2]
    standard_mass_flow=(SLPM*rho_std)/60 # from SLPM to standard mass flow rate [g/s]
    return standard_mass_flow