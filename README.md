# BOMshot

## Summary
BOMshot is a python script for Fusion 360 to generate an Excel Spreadsheet Bill of Materials with component image captures and optional .STEP file export

The spreadsheet generated includes a cover sheet with project details and company logo that is imported from the script directory resources folder

Exported files (images and STEP) are stored hierarchically by component assembly

## Script Dialog

![image](https://github.com/Veldraeos/bomshot/assets/75970270/afe466c7-e6c3-42ce-be37-382871206018)

## Cover Sheet

![image](https://github.com/Veldraeos/bomshot/assets/75970270/080aace7-afcf-4b7b-b442-a65285fbe04e)

## BOM

![image](https://github.com/Veldraeos/bomshot/assets/75970270/f7d0def0-5796-458d-8b1a-583390dbea3f)


## Installation
1. Clone this repository to your local machine
2. In Fusion 360 navigate to UTILITIES tab, then click ADD-INS toolbar button
   
   ![image](https://github.com/Veldraeos/bomshot/assets/75970270/c97c8425-a9da-4a00-ae37-5e79e6e418bb)
4. Click + symbol next to My Scripts
   
   ![image](https://github.com/Veldraeos/bomshot/assets/75970270/9abc6385-ef17-4261-9cff-b68d2e1a8068)
6. Navigate to the directory containing BOMshot.py (where you cloned this repository) and click Select Folder
   
   ![image](https://github.com/Veldraeos/bomshot/assets/75970270/670e8866-43c2-4f76-ac10-eb8e3153efb2)
8. You should now see the script under My Scripts

## Execution
1. In Fusion 360 navigate to UTILITIES tab, then click ADD-INS toolbar button
   
   ![image](https://github.com/Veldraeos/bomshot/assets/75970270/c97c8425-a9da-4a00-ae37-5e79e6e418bb)
3. Select BOMshot from My Scripts section
4. Click Run
5. Input relevant project details and select whether to export .STEP files
6. Click OK
7. Select output directory
8. Wait for the script to complete
9. You will then be prompted to open the BOM spreadsheet
    
    ![image](https://github.com/Veldraeos/bomshot/assets/75970270/5f497343-5a4a-42ee-ab06-6914ed431118)
