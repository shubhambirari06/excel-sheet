export const ecm2InputVariables = [
    { 'Input Variables': 'System Efficiency (Heating)', 'Value': 'See Below' },
    { 'Input Variables': 'System Efficiency (Cooling EER)', 'Value': 'See Below' },
    { 'Input Variables': 'Cost of Electricity ($/ kwh)', 'Value': 0.12 },
    { 'Input Variables': 'Cost of Electricity ($/ Therm)', 'Value': 1.3 },
];

export const ecm2Equations = [
    { 'Equations & Variables': 'Energy Equation (Cooling kWh)', 'Formula': '=Tons*(12/EER)*Bin Hours' },
    { 'Equations & Variables': 'Energy Equation (Heating)', 'Formula': '=(1.08*CFM*DT*Bin Hours)/3412.14' },
    { 'Equations & Variables': 'Fan Savings (kWh)', 'Formula': '=((HP*Mtr Load %*.746)/Mtr Eff * VFD eff) * VFD Speed^fan law exponent' },
];

export const ecm2EquipmentSummary = [
    { 'Parameter': 'Cooling Tons', 'AHU #1': 39.10, 'AHU #2': 61.60, 'AHU #3': 66.50, 'AHU #4': 17.60, 'AHU #5': 2.60, 'AHU #6': 7.50, 'AHU #7': 3.40, 'RTU # 1': 15.00 },
    { 'Parameter': 'Heating MBH', 'AHU #1': 469.20, 'AHU #2': 739.20, 'AHU #3': 798.00, 'AHU #4': 211.20, 'AHU #5': 31.20, 'AHU #6': 90.00, 'AHU #7': 40.80, 'RTU # 1': 180.00 },
    { 'Parameter': 'Heating eff %', 'AHU #1': 0.85, 'AHU #2': 0.85, 'AHU #3': 0.85, 'AHU #4': 0.85, 'AHU #5': 0.85, 'AHU #6': 0.85, 'AHU #7': 0.85, 'RTU # 1': 0.85 },
    { 'Parameter': 'Cooling EER', 'AHU #1': 11.3, 'AHU #2': 11.3, 'AHU #3': 11.3, 'AHU #4': 11.3, 'AHU #5': 11.3, 'AHU #6': 11.3, 'AHU #7': 11.3, 'RTU # 1': 10.0 },
    { 'Parameter': 'Cooling Eff (kw/Ton)', 'AHU #1': 1.1, 'AHU #2': 1.1, 'AHU #3': 1.1, 'AHU #4': 1.1, 'AHU #5': 1.1, 'AHU #6': 1.1, 'AHU #7': 1.1, 'RTU # 1': 1.2 },
    { 'Parameter': 'Supply Fan Max CFM', 'AHU #1': 15640, 'AHU #2': 24640, 'AHU #3': 26600, 'AHU #4': 7040, 'AHU #5': 1040, 'AHU #6': 3000, 'AHU #7': 1360, 'RTU # 1': 6000 },
    { 'Parameter': 'Supply Fan Min CFM', 'AHU #1': 3910, 'AHU #2': 6160, 'AHU #3': 6650, 'AHU #4': 1760, 'AHU #5': 260, 'AHU #6': 750, 'AHU #7': 340, 'RTU # 1': 1500 },
];

export const ecm2EnergySavings = [
    { 'Equipment': 'AHU #1', 'Heating Energy Savings (Therms)': 653, ' ': null, 'Cost Savings ($)': 849.40 },
    { 'Equipment': 'AHU #2', 'Heating Energy Savings (Therms)': 20, ' ': null, 'Cost Savings ($)': 25.83 },
    { 'Equipment': 'AHU #3', 'Heating Energy Savings (Therms)': 185, ' ': null, 'Cost Savings ($)': 240.87 },
    { 'Equipment': 'AHU #4', 'Heating Energy Savings (Therms)': 139, ' ': null, 'Cost Savings ($)': 180.48 },
    { 'Equipment': 'AHU #5', 'Heating Energy Savings (Therms)': 16, ' ': null, 'Cost Savings ($)': 20.59 },
    { 'Equipment': 'AHU #6', 'Heating Energy Savings (Therms)': 2, ' ': null, 'Cost Savings ($)': 2.30 },
    { 'Equipment': 'AHU #7', 'Heating Energy Savings (Therms)': 0.0, ' ': null, 'Cost Savings ($)': 0.05 },
    { 'Equipment': 'RTU # 1', 'Heating Energy Savings (Therms)': 129, ' ': null, 'Cost Savings ($)': 167.51 },
    { 'Equipment': 'Totals', 'Heating Energy Savings (Therms)': 1144, ' ': null, 'Cost Savings ($)': 1487.04 },
];

export const ecm2TotalGasSavings = [
    { 'Total Gas Savings (Heating)': 'kHw', 'Value': 1143.87 },
    { 'Total Gas Savings (Heating)': 'Dollars', 'Value': 1487.04 },
];

export const ecm2RtuLoadProfile = [
    { 'OAT': 102.5, 'EFLCH %': '100%', 'EFLHH %': '' },
    { 'OAT': 55, 'EFLCH %': '0%', 'EFLHH %': '' },
    { 'OAT': 55, 'EFLCH %': '', 'EFLHH %': '0%' },
    { 'OAT': -2.5, 'EFLCH %': '', 'EFLHH %': '100%' },
    { 'OAT': 'Slope', 'EFLCH %': 0.021052632, 'EFLHH %': -0.017391304 },
    { 'OAT': 'Intercept', 'EFLCH %': -1.157894737, 'EFLHH %': 0.956521739 },
];

const createBinRow = (temp: number | null, hours: number | null, coolingLoad: number | null, eflh: string | null, cfm: number | null, satEx: number | null, satProp: number | null, dtProp: number | null, thermsSaved: number | null) => ({
    'Temp Mid-Point': temp,
    'Bin Hours Proposed': hours,
    'Cooling Load (MBH)': coolingLoad,
    'Equivalent Full Load Hours (%) at Each Bin': eflh,
    'Supply Fan CFM': cfm,
    'Supply Air Temperature (Existing)': satEx,
    'Supply Air Temperature (Proposed)': satProp,
    ' ': null,
    'Htg dT Proposed (deg F)': dtProp,
    'Heating Therms Saved': thermsSaved
});

export const ecm2BinDataAHU1 = [
    createBinRow(102.5, 3, null, '100.00%', null, null, null, null, null),
    createBinRow(97.5, 28, null, '89.47%', null, null, null, null, null),
    createBinRow(92.5, 72, null, '78.95%', null, null, null, null, null),
    createBinRow(87.5, 192, null, '68.42%', null, null, null, null, null),
    createBinRow(82.5, 420, null, '57.89%', null, null, null, null, null),
    createBinRow(77.5, 649, null, '47.37%', null, null, null, null, null),
    createBinRow(72.5, 819, null, '36.84%', null, null, null, null, null),
    createBinRow(67.5, 873, null, '26.32%', null, null, null, null, null),
    createBinRow(62.5, 782, null, '15.79%', null, null, null, null, null),
    createBinRow(57.5, 768, null, '5.26%', null, null, null, null, null),
    createBinRow(52.5, 795, 20.40, '4.35%', 7448, 70, 75, 5.0, 13.90114286),
    createBinRow(47.5, 551, 61.20, '13.04%', 8192, 70, 75, 5.0, 31.79427429),
    createBinRow(42.5, 858, 102.00, '21.74%', 8937, 70, 75, 5.0, 90.01645714),
    createBinRow(37.5, 651, 142.80, '30.43%', 9682, 70, 75, 5.0, 103.58712),
    createBinRow(32.5, 476, 183.60, '39.13%', 10427, 70, 75, 5.0, 104.87232),
    createBinRow(27.5, 283, 224.40, '47.83%', 11171, 70, 75, 5.0, 81.64954286),
    createBinRow(22.5, 277, 265.20, '56.52%', 11916, 70, 75, 5.0, 100.7456914),
    createBinRow(17.5, 178, 306.00, '65.22%', 12661, 70, 75, 5.0, 79.36765714),
    createBinRow(12.5, 71, 346.80, '73.91%', 13406, 70, 75, 5.0, 37.98946286),
    createBinRow(7.5, 9, 387.60, '82.61%', 14150, 70, 75, 5.0, 5.681108571),
    createBinRow(2.5, 4, 428.40, '91.30%', 14895, 70, 75, 5.0, 2.9376),
    createBinRow(-2.5, 1, 469.20, '100.00%', 15640, 70, 75, 5.0, 0.84456),
    createBinRow(null, null, null, null, null, null, null, null, 653.39)
];

export const ecm2BinDataAHU2 = [
    createBinRow(102.5, 2, null, '100.00%', 65, 70, null, null, null),
    createBinRow(97.5, 2, null, '89.47%', 65, 70, null, null, null),
    createBinRow(92.5, 6, null, '78.95%', 65, 70, null, null, null),
    createBinRow(87.5, 20, null, '68.42%', 65, 70, null, null, null),
    createBinRow(82.5, 37, null, '57.89%', 65, 70, null, null, null),
    createBinRow(77.5, 20, null, '47.37%', 65, 70, null, null, null),
    createBinRow(72.5, 23, null, '36.84%', 65, 70, null, null, null),
    createBinRow(67.5, 14, null, '26.32%', 65, 70, null, null, null),
    createBinRow(62.5, 15, null, '15.79%', 65, 70, null, null, null),
    createBinRow(57.5, 30, null, '5.26%', 65, 70, null, null, null),
    createBinRow(52.5, 21, 32.14, '4.35%', 11733, 70, 75, 5.0, 0.578504348),
    createBinRow(47.5, 24, 96.42, '13.04%', 12907, 70, 75, 5.0, 2.181787826),
    createBinRow(42.5, 11, 160.70, '21.74%', 14080, 70, 75, 5.0, 1.818156522),
    createBinRow(37.5, 11, 224.97, '30.43%', 15253, 70, 75, 5.0, 2.757537391),
    createBinRow(32.5, 3, 289.25, '39.13%', 16427, 70, 75, 5.0, 1.041307826),
    createBinRow(27.5, 10, 353.53, '47.83%', 17600, 70, 75, 5.0, 4.545391304),
    createBinRow(22.5, 6, 417.81, '56.52%', 18773, 70, 75, 5.0, 3.437968696),
    createBinRow(17.5, 5, 482.09, '65.22%', 19947, 70, 75, 5.0, 3.512347826),
    createBinRow(12.5, 0, 546.37, '73.91%', 21120, 70, 75, 5.0, 0),
    createBinRow(7.5, 0, 610.64, '82.61%', 22293, 70, 75, 5.0, 0),
    createBinRow(2.5, 0, 674.92, '91.30%', 23467, 70, 75, 5.0, 0),
    createBinRow(-2.5, 0, 739.20, '100.00%', 24640, 70, 75, 5.0, 0),
    createBinRow(null, null, null, null, null, null, null, null, 19.87)
];

export const ecm2BinDataAHU3 = [
    createBinRow(102.5, 0, null, '100.00%', 65, 70, null, null, null),
    createBinRow(97.5, 0, null, '89.47%', 65, 70, null, null, null),
    createBinRow(92.5, 2, null, '78.95%', 65, 70, null, null, null),
    createBinRow(87.5, 11, null, '68.42%', 65, 70, null, null, null),
    createBinRow(82.5, 77, null, '57.89%', 65, 70, null, null, null),
    createBinRow(77.5, 187, null, '47.37%', 65, 70, null, null, null),
    createBinRow(72.5, 211, null, '36.84%', 65, 70, null, null, null),
    createBinRow(67.5, 203, null, '26.32%', 65, 70, null, null, null),
    createBinRow(62.5, 158, null, '15.79%', 65, 70, null, null, null),
    createBinRow(57.5, 153, null, '5.26%', 65, 70, null, null, null),
    createBinRow(52.5, 145, 34.70, '4.35%', 12667, 70, 75, 5.0, 4.312173913),
    createBinRow(47.5, 206, 104.09, '13.04%', 13933, 70, 75, 5.0, 20.21666087),
    createBinRow(42.5, 162, 173.48, '21.74%', 15200, 70, 75, 5.0, 28.90643478),
    createBinRow(37.5, 103, 242.87, '30.43%', 16467, 70, 75, 5.0, 27.87448696),
    createBinRow(32.5, 79, 312.26, '39.13%', 17733, 70, 75, 5.0, 29.60233043),
    createBinRow(27.5, 65, 381.65, '47.83%', 19000, 70, 75, 5.0, 31.89521739),
    createBinRow(22.5, 39, 451.04, '56.52%', 20267, 70, 75, 5.0, 24.12438261),
    createBinRow(17.5, 23, 520.43, '65.22%', 21533, 70, 75, 5.0, 17.442),
    createBinRow(12.5, 1, 589.83, '73.91%', 22800, 70, 75, 5.0, 0.910017391),
    createBinRow(7.5, 0, 659.22, '82.61%', 24067, 70, 75, 5.0, 0),
    createBinRow(2.5, 0, 728.61, '91.30%', 25333, 70, 75, 5.0, 0),
    createBinRow(-2.5, 0, 798.00, '100.00%', 26600, 70, 75, 5.0, 0),
    createBinRow(null, null, null, null, null, null, null, null, 185.28)
];

export const ecm2BinDataAHU4 = [
    createBinRow(102.5, 5, null, '100.00%', 65, 70, null, null, null),
    createBinRow(97.5, 40, null, '89.47%', 65, 70, null, null, null),
    createBinRow(92.5, 229, null, '78.95%', 65, 70, null, null, null),
    createBinRow(87.5, 435, null, '68.42%', 65, 70, null, null, null),
    createBinRow(82.5, 713, null, '57.89%', 65, 70, null, null, null),
    createBinRow(77.5, 775, null, '47.37%', 65, 70, null, null, null),
    createBinRow(72.5, 1091, null, '36.84%', 65, 70, null, null, null),
    createBinRow(67.5, 1039, null, '26.32%', 65, 70, null, null, null),
    createBinRow(62.5, 894, null, '15.79%', 65, 70, null, null, null),
    createBinRow(57.5, 759, null, '5.26%', 65, 70, null, null, null),
    createBinRow(52.5, 640, 9.18, '4.35%', 3352, 70, 75, 5.0, 5.03731677),
    createBinRow(47.5, 671, 27.55, '13.04%', 3688, 70, 75, 5.0, 17.42832894),
    createBinRow(42.5, 673, 45.91, '21.74%', 4023, 70, 75, 5.0, 31.7823205),
    createBinRow(37.5, 287, 64.28, '30.43%', 4358, 70, 75, 5.0, 20.55618783),
    createBinRow(32.5, 196, 82.64, '39.13%', 4693, 70, 75, 5.0, 19.43774609),
    createBinRow(27.5, 214, 101.01, '47.83%', 5029, 70, 75, 5.0, 27.79182112),
    createBinRow(22.5, 83, 119.37, '56.52%', 5364, 70, 75, 5.0, 13.58816199),
    createBinRow(17.5, 10, 137.74, '65.22%', 5699, 70, 75, 5.0, 2.007055901),
    createBinRow(12.5, 5, 156.10, '73.91%', 6034, 70, 75, 5.0, 1.20423354),
    createBinRow(7.5, 1, 174.47, '82.61%', 6370, 70, 75, 5.0, 0.284136149),
    createBinRow(2.5, 0, 192.83, '91.30%', 6705, 70, 75, 5.0, 0),
    createBinRow(-2.5, 0, 211.20, '100.00%', 7040, 70, 75, 5.0, 0),
    createBinRow(null, null, null, null, null, null, null, null, 138.83)
];

export const ecm2BinDataAHU5 = [
    createBinRow(102.5, 3, null, '100.00%', 65, 70, null, null, null),
    createBinRow(97.5, 26, null, '89.47%', 65, 70, null, null, null),
    createBinRow(92.5, 67, null, '78.95%', 65, 70, null, null, null),
    createBinRow(87.5, 172, null, '68.42%', 65, 70, null, null, null),
    createBinRow(82.5, 356, null, '57.89%', 65, 70, null, null, null),
    createBinRow(77.5, 440, null, '47.37%', 65, 70, null, null, null),
    createBinRow(72.5, 344, null, '36.84%', 65, 70, null, null, null),
    createBinRow(67.5, 325, null, '26.32%', 65, 70, null, null, null),
    createBinRow(62.5, 334, null, '15.79%', 65, 70, null, null, null),
    createBinRow(57.5, 346, null, '5.26%', 65, 70, null, null, null),
    createBinRow(52.5, 337, 1.36, '4.35%', 495, 70, 75, 5.0, 0.391840994),
    createBinRow(47.5, 229, 4.07, '13.04%', 545, 70, 75, 5.0, 0.878677267),
    createBinRow(42.5, 328, 6.78, '21.74%', 594, 70, 75, 5.0, 2.288258385),
    createBinRow(37.5, 227, 9.50, '30.43%', 644, 70, 75, 5.0, 2.401857391),
    createBinRow(32.5, 193, 12.21, '39.13%', 693, 70, 75, 5.0, 2.827533913),
    createBinRow(27.5, 93, 14.92, '47.83%', 743, 70, 75, 5.0, 1.784213665),
    createBinRow(22.5, 108, 17.63, '56.52%', 792, 70, 75, 5.0, 2.61196323),
    createBinRow(17.5, 65, 20.35, '65.22%', 842, 70, 75, 5.0, 1.927229814),
    createBinRow(12.5, 18, 23.06, '73.91%', 891, 70, 75, 5.0, 0.640433292),
    createBinRow(7.5, 2, 25.77, '82.61%', 941, 70, 75, 5.0, 0.083949317),
    createBinRow(2.5, 1, 28.49, '91.30%', 990, 70, 75, 5.0, 0.048834783),
    createBinRow(-2.5, 1, 31.20, '100.00%', 1040, 70, 75, 5.0, 0.05616),
    createBinRow(null, null, null, null, null, null, null, null, 15.84)
];

export const ecm2BinDataAHU6 = [
    createBinRow(102.5, 3, null, '100.00%', 65, 70, null, null, null),
    createBinRow(97.5, 5, null, '89.47%', 65, 70, null, null, null),
    createBinRow(92.5, 20, null, '78.95%', 65, 70, null, null, null),
    createBinRow(87.5, 39, null, '68.42%', 65, 70, null, null, null),
    createBinRow(82.5, 48, null, '57.89%', 65, 70, null, null, null),
    createBinRow(77.5, 41, null, '47.37%', 65, 70, null, null, null),
    createBinRow(72.5, 30, null, '36.84%', 65, 70, null, null, null),
    createBinRow(67.5, 21, null, '26.32%', 65, 70, null, null, null),
    createBinRow(62.5, 39, null, '15.79%', 65, 70, null, null, null),
    createBinRow(57.5, 42, null, '5.26%', 65, 70, null, null, null),
    createBinRow(52.5, 23, 3.91, '4.35%', 1429, 70, 75, 5.0, 0.0),
    createBinRow(47.5, 37, 11.74, '13.04%', 1571, 70, 75, 5.0, 0.1),
    createBinRow(42.5, 18, 19.57, '21.74%', 1714, 70, 75, 5.0, 0.1),
    createBinRow(37.5, 14, 27.39, '30.43%', 1857, 70, 75, 5.0, 0.2),
    createBinRow(32.5, 6, 35.22, '39.13%', 2000, 70, 75, 5.0, 0.1),
    createBinRow(27.5, 16, 43.04, '47.83%', 2143, 70, 75, 5.0, 0.5),
    createBinRow(22.5, 9, 50.87, '56.52%', 2286, 70, 75, 5.0, 0.4),
    createBinRow(17.5, 5, 58.70, '65.22%', 2429, 70, 75, 5.0, 0.3),
    createBinRow(12.5, 0, 66.52, '73.91%', 2571, 70, 75, 5.0, 0.0),
    createBinRow(7.5, 0, 74.35, '82.61%', 2714, 70, 75, 5.0, 0.0),
    createBinRow(2.5, 0, 82.17, '91.30%', 2857, 70, 75, 5.0, 0.0),
    createBinRow(-2.5, 0, 90.00, '100.00%', 3000, 70, 75, 5.0, 0.0),
    createBinRow(null, null, null, null, null, null, null, null, 1.77)
];

export const ecm2BinDataAHU7 = [
    createBinRow(102.5, 3, null, '100.00%', 65, 70, null, null, null),
    createBinRow(97.5, 26, null, '89.47%', 65, 70, null, null, null),
    createBinRow(92.5, 67, null, '78.95%', 65, 70, null, null, null),
    createBinRow(87.5, 172, null, '68.42%', 65, 70, null, null, null),
    createBinRow(82.5, 356, null, '57.89%', 65, 70, null, null, null),
    createBinRow(77.5, 440, null, '47.37%', 65, 70, null, null, null),
    createBinRow(72.5, 344, null, '36.84%', 65, 70, null, null, null),
    createBinRow(67.5, 325, null, '26.32%', 65, 70, null, null, null),
    createBinRow(62.5, 334, null, '15.79%', 65, 70, null, null, null),
    createBinRow(57.5, 346, null, '5.26%', 65, 70, null, null, null),
    createBinRow(52.5, 337, 1.77, '4.35%', 648, 70, 75, 5.0, 0.0),
    createBinRow(47.5, 229, 5.32, '13.04%', 712, 70, 75, 5.0, 0.3),
    createBinRow(42.5, 328, 8.87, '21.74%', 777, 70, 75, 5.0, 1.0),
    createBinRow(37.5, 227, 12.42, '30.43%', 842, 70, 75, 5.0, 1.4),
    createBinRow(32.5, 193, 15.97, '39.13%', 907, 70, 75, 5.0, 1.9),
    createBinRow(27.5, 93, 19.51, '47.83%', 971, 70, 75, 5.0, 1.4),
    createBinRow(22.5, 108, 23.06, '56.52%', 1036, 70, 75, 5.0, 2.2),
    createBinRow(17.5, 65, 26.61, '65.22%', 1101, 70, 75, 5.0, 1.8),
    createBinRow(12.5, 18, 30.16, '73.91%', 1166, 70, 75, 5.0, 0.6),
    createBinRow(7.5, 2, 33.70, '82.61%', 1230, 70, 75, 5.0, 0.1),
    createBinRow(2.5, 1, 37.25, '91.30%', 1295, 70, 75, 5.0, 0.1),
    createBinRow(-2.5, 1, 40.80, '100.00%', 1360, 70, 75, 5.0, 0.1),
    createBinRow(null, null, null, null, null, null, null, null, 0.04)
];

export const ecm2BinDataRTU1 = [
    createBinRow(102.5, 3, null, '100.00%', 65, 70, null, null, null),
    createBinRow(97.5, 28, null, '89.47%', 65, 70, null, null, null),
    createBinRow(92.5, 72, null, '78.95%', 65, 70, null, null, null),
    createBinRow(87.5, 192, null, '68.42%', 65, 70, null, null, null),
    createBinRow(82.5, 420, null, '57.89%', 65, 70, null, null, null),
    createBinRow(77.5, 649, null, '47.37%', 65, 70, null, null, null),
    createBinRow(72.5, 819, null, '36.84%', 65, 70, null, null, null),
    createBinRow(67.5, 873, null, '26.32%', 65, 70, null, null, null),
    createBinRow(62.5, 782, null, '15.79%', 65, 70, null, null, null),
    createBinRow(57.5, 768, null, '5.26%', 65, 70, null, null, null),
    createBinRow(52.5, 795, 7.83, '4.35%', 2857, 70, 75, 5.0, 0.4),
    createBinRow(47.5, 551, 23.48, '13.04%', 3143, 70, 75, 5.0, 2.7),
    createBinRow(42.5, 858, 39.13, '21.74%', 3429, 70, 75, 5.0, 11.6),
    createBinRow(37.5, 651, 54.78, '30.43%', 3714, 70, 75, 5.0, 17.2),
    createBinRow(32.5, 476, 70.43, '39.13%', 4000, 70, 75, 5.0, 20.8),
    createBinRow(27.5, 283, 86.09, '47.83%', 4286, 70, 75, 5.0, 18.4),
    createBinRow(22.5, 277, 101.74, '56.52%', 4571, 70, 75, 5.0, 25.2),
    createBinRow(17.5, 178, 117.39, '65.22%', 4857, 70, 75, 5.0, 21.6),
    createBinRow(12.5, 71, 133.04, '73.91%', 5143, 70, 75, 5.0, 11.0),
    createBinRow(7.5, 9, 148.70, '82.61%', 5429, 70, 75, 5.0, 1.7),
    createBinRow(2.5, 4, 164.35, '91.30%', 5714, 70, 75, 5.0, 0.9),
    createBinRow(-2.5, 1, 180.00, '100.00%', 6000, 70, 75, 5.0, 0.3),
    createBinRow(null, null, null, null, null, null, null, null, 128.85)
];