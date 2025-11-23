# SAP Outbound Shipping Automation

<p align="center">
  <img src="automation-tool.png" width="1000" title="SAP Shipping Tool Interface">
</p>

## Context & Impact
I developed this automation tool for the outbound shipping team at a Fortune 200 manufacturing company. Our manual process in SAP GUI was repetitive and prone to minor data entry errors.

Originally built to streamline my own workflow, the tool was subsequently reviewed by the Lead Logistics Manager and **adopted by the wider shipping department.** It now serves as the standard utility for processing outbound deliveries, ensuring consistent documentation and faster processing times across the team.

## What It Actually Does
The tool acts as a "force multiplier" for the shipping team. Instead of manually navigating transaction code `VL02N` (Change Outbound Delivery) for every single order, the script:

1.  **Standardizes Input:** A unified GUI prompts the user for Delivery Number, weights, dims, and document counts, ensuring data consistency across all analysts.
2.  **Automates Logic:**
    * **Plant Determination:** It automatically detects if the order is from Plant 1812 or 1814 and selects the strictly required SAP Output Type (`ZPL0` vs `YPLA`).
    * **Smart Routing:** It dynamically switches between "Full Documentation," "Packing List Only," or "BOL Only" modes based on the user's input, skipping unnecessary SAP screens to save time.
3.  **Executes:** It handles the tedious creation of Handling Units (HU), assigns physical specs, and triggers the print jobs.
4.  **Handoff:** It exports PDF paperwork to a standardized OneDrive path and hands control back to the user for the final "Post Goods Issue" (PGI), maintaining a "Human-in-the-Loop" workflow for safety and validation.

## Technical Challenges (Production-Ready Code)
Because this tool is used by multiple team members in a live production environment, it had to be crash-proof. I implemented several resilience features to handle the unpredictability of SAP:

### 1. Handling "Random" SAP Pop-ups
Different materials trigger different warnings in SAP (e.g., "Missing Country of Origin" or "Serial Number" checks). I wrote logic that actively monitors the SAP Status Bar and dismisses these pop-ups automatically so the automation doesn't hang on a user's screen.

### 2. Memory Leaks & Garbage Collection
SAP GUI sessions often degrade over time during high-volume processing. I implemented a `CleanupMemory` subroutine that forces garbage collection and, if necessary, auto-refreshes the SAP session after a set number of transactions to prevent memory leaks during long shifts.

### 3. Connection "Healing"
If a user's SAP session drops or freezes, the script detects the broken object reference and attempts to reconnect to the active session automatically, minimizing downtime and frustration for the user.

## Tech Stack
* **Language:** VBScript (Native language for SAP GUI Scripting)
* **Interface:** SAP GUI Scripting API
* **Deployment:** Shared Internal Tool

---
*Disclaimer: This code is tailored to a specific SAP environment configuration (transaction flows, screen IDs, and plant codes). It serves as a demonstration of business process automation and error handling techniques.*
