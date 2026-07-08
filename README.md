# Kuldeep Singh Resume Format

Version-controlled resume repository for Kuldeep Singh.

## Current Version

- `R01`: EV ECU Software Lead / MATLAB-Simulink Electrification resume
- `R02`: Targeted EV/software resume pack for Tata/Tata Technologies, Tesla, Daimler/Mercedes, BMW, JLR/Land Rover, AVL, Bosch, and Continental
- `R03`: Generic 2-page EV Software Lead / Technical Manager resume for broad lead-role applications

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
- `resumes/R03/*.docx`: editable 2-page generic lead resume
- `resumes/R03/*.pdf`: rendered PDF copy for broad applications
- `resumes/R03/*.txt`: ATS/plain-text copy
- `resumes/R03/README.md`: R03 package notes
- `source/build_ev_resume.py`: generator script used to create the resume package
- `source/create_r02_targeted_resume_pack.py`: generator script used to create the R02 targeted resume pack
- `source/create_r03_generic_lead_resume.py`: generator script used to create the R03 generic lead resume
