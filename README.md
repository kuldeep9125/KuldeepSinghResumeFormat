# Kuldeep Singh Resume Format

Version-controlled resume repository for Kuldeep Singh.

## Current Version

- `R01`: EV ECU Software Lead / MATLAB-Simulink Electrification resume
- `R02`: Targeted EV/software resume pack for Tata/Tata Technologies, Tesla, Daimler/Mercedes, BMW, JLR/Land Rover, AVL, Bosch, and Continental
- `R03`: Generic 2-page EV Software Lead / Technical Manager resume for broad lead-role applications
- `R04`: Automotive-benchmarked EV Software & Technical Lead resume optimized against current industry screening signals
- `R05`: Project-verified full two-page EV Software & Technical Lead resume integrating Simulink/Stateflow evidence
- `R06`: Privacy-clean full two-page EV Software & Technical Lead resume with project/file identifiers removed from resume text
- `R07`: HR-reviewed duplicate-reduced EV Technical Manager resume with cleaner two-page visual balance

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
- `resumes/R04/*.docx`: editable automotive-benchmarked lead resume
- `resumes/R04/*.pdf`: rendered PDF copy for automotive lead applications
- `resumes/R04/*.txt`: ATS/plain-text copy
- `resumes/R04/R04_Benchmark_Notes.md`: comparison notes and research-backed improvement rationale
- `resumes/R05/*.docx`: editable project-verified full two-page resume
- `resumes/R05/*.pdf`: rendered PDF copy for applications
- `resumes/R05/*.txt`: ATS/plain-text copy
- `resumes/R05/R05_Project_Scan_Notes.md`: high-level source archive scan and integration notes
- `resumes/R06/*.docx`: editable privacy-clean full two-page resume
- `resumes/R06/*.pdf`: rendered PDF copy for applications
- `resumes/R06/*.txt`: ATS/plain-text copy
- `resumes/R06/R06_Privacy_Clean_Project_Notes.md`: high-level privacy-clean project notes
- `resumes/R07/*.docx`: editable HR-reviewed technical-manager resume
- `resumes/R07/*.pdf`: rendered PDF copy for applications
- `resumes/R07/*.txt`: ATS/plain-text copy
- `resumes/R07/R07_HR_Review_Notes.md`: HR review notes and duplicate-reduction rationale
- `source/build_ev_resume.py`: generator script used to create the resume package
- `source/create_r02_targeted_resume_pack.py`: generator script used to create the R02 targeted resume pack
- `source/create_r03_generic_lead_resume.py`: generator script used to create the R03 generic lead resume
- `source/create_r04_automotive_benchmarked_resume.py`: generator script used to create the R04 benchmarked resume
- `source/create_r05_project_verified_full_resume.py`: generator script used to create the R05 project-verified resume
- `source/create_r06_privacy_clean_full_resume.py`: generator script used to create the R06 privacy-clean resume
- `source/create_r07_hr_reviewed_technical_manager_resume.py`: generator script used to create the R07 HR-reviewed resume
