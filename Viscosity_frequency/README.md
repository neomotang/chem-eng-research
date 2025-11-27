# Determining viscosity of mixtures from the resonant frequency of a quartz crystal

---

## Overview
In this project, the resonant frequency of a quartz crystal in contact with fluid mixtures was used to determine the viscosity of those mixtures. The resonant frequency was determined as the frequency at which an induced wave into the crystal resulted in a measured response voltage of minimum amplitude. The PicoScope 2206B USB oscilloscope was used to generate induced waves at a range of frequencies and also measure the crystal's response. Measured responses (as root mean square voltage) were then approximated as quadratic functions in order to find the frequency corresponding to the minimum voltage. The frequency was used along with the densities of the mixtures to determine the mixture viscosity. Results were published in a journal article (Motang et al., 2023) (https://doi.org/10.1016/j.supflu.2023.105864).

## Tools Used
- `Python`: Used to control the PicoScope (signal generator and oscilloscope) and also save measured data in a .csv file.
- `Excel VBA`: Used to determine the resonant frequency from the .csv files, approximating the voltage-frequency curve as a quadratic function.
