# R05 Project Scan Notes

## Source Reviewed

Local archive: `/Users/kuldeepsingh/Downloads/HybridProjectpatent/NIDECG02+MICROVAST.zip`.

The archive has damaged/incomplete central-directory metadata, but streaming extraction and local-header scans showed a real NIDECG02 + MICROVAST project folder with Simulink, MATLAB, DBC, MF4, generated target, calibration, and log artifacts.

## Resume-Safe Evidence Extracted

- Final Simulink model revisions: `G02_MICROVASTR03_FINAL.slx` and `G02_MICROVASTR10_FINAL.slx`.
- Simulink metadata: R2018a model files with `HybridControlStrategy` harness entries, Stateflow XML, block-diagram XML, model workspace, and code dictionary artifacts.
- MATLAB communication/control files: `VCVCCU_J1939.m`, `TorqueSpeedCtrl_LE_CAN.m`, `BMS_Charge_New.m`, `IXMV28.m`, `ZETTAJOULELTOBATTERY.m`, and `telematics.m`.
- DBC evidence: `TATA1618_TSRTC.dbc` with 26 CAN messages and 211 signals across BMS, VCU command, motor status, telematics, EV faults, remote access, charge, precharge, voltage/current, torque/speed, and diagnostics.
- Build/calibration evidence: generated `.mot` firmware outputs and `.a2l` measurement/calibration files under `Target_out`.
- Test and calibration evidence: `.mf4` logs and code-log sheets showing assist/regen, temperature derate, cut-off voltage, and calibration update notes.

## R05 Resume Integration

- Added a page-1 project-verified proof section to reduce blank page space and make the resume more credible for automotive software lead screening.
- Integrated only high-level, resume-safe facts; no source code, CAN IDs, proprietary logic, or copied implementation details are included in the resume.
- Rebalanced R04's early forced page break into a fuller two-page layout.

## Prior Benchmark Sources Retained

- https://www.tesla.com/careers/search
- https://dtici.daimlertruck.com/career/
- https://www.tatatechnologies.com/us/solutions/
- https://www.jaguarlandrovercareers.com/
- https://www.avl.com/en/career/fields-interest/e-mobility
- https://www.bosch.com/careers/
- https://www.continental.com/en/career/
