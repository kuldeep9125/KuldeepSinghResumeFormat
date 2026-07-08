# Kuldeep Singh Resume Format

Version-controlled resume repository for Kuldeep Singh.

## Current Version

- `R01`: EV ECU Software Lead / MATLAB-Simulink Electrification resume
- `R02`: Targeted EV/software resume pack for Tata/Tata Technologies, Tesla, Daimler/Mercedes, BMW, JLR/Land Rover, AVL, Bosch, and Continental

## Versioning Rule

Every resume update should create a new versioned file set:

- `R01`, `R02`, `R03`, etc.
- Keep the version in the resume filename.
- Do not overwrite older released versions.

Example:

```text
resumes/R02/Kuldeep_Singh_EV_ECU_Software_Lead_ATS_Resume_R02.docx
resumes/R02/Kuldeep_Singh_EV_ECU_Software_Lead_ATS_Resume_R02.pdf
resumes/R02/Kuldeep_Singh_EV_ECU_Software_Lead_ATS_Resume_R02.txt
```

## Included Files

- `resumes/R01/*.docx`: editable Word resume
- `resumes/R01/*.pdf`: rendered PDF copy for quick sharing
- `resumes/R01/*.txt`: ATS/plain-text copy
- `resumes/R02/*/*.docx`: targeted editable Word resumes
- `resumes/R02/*/*.pdf`: rendered PDF copies for quick sharing
- `resumes/R02/*/*.txt`: ATS/plain-text copies
- `resumes/R02/*/WHERE_TO_APPLY.md`: official application links and targeting notes
- `source/build_ev_resume.py`: generator script used to create the resume package
- `source/create_r02_targeted_resume_pack.py`: generator script used to create the R02 targeted resume pack
