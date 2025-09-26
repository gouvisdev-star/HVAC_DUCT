;; HVAC_DUCT - Simple Auto Load
;; Load DLL and enable auto-loading

;; Load the DLL
(arxload "D:\\PROJECT_API_CAD\\MECH\\HVAC_DUCT\\HVAC_DUCT\\HVAC_DUCT\\bin\\Debug\\HVAC_DUCT.dll")

;; Enable auto-cleanup
(command "REGISTERERASE")

;; Load ticks
(command "AUTO")

;; Show success message
(if (member "HVAC_DUCT" (arx))
  (princ "\nHVAC_DUCT loaded and ready!")
  (princ "\nHVAC_DUCT failed to load!")
)

(princ)
