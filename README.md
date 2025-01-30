# OutlookVBA
**Copy-PST.vba**
 'This script copies all items from the source PST files in a folder to the destination PST files. 
' why? The Enterprise Vault exports PST and has all item marked as archived and the attachments are converted as a link. 
' This script to make the item as a regular item again so they can be imported to Exchange online.

**Delete-EnterpriseVault-tagged-items.vba**
  'This script to delete the Enterprise Vault tagged items on a given folder if you accidently import the PST exported by Enterprise Vault without un-vault them.
