

Function ATTicketSwitch {
  param ($ticketStatus)
    
    switch ($ticketStatus) {
        1 {"New"}
        5 {"Complete"}
        7 {"Waiting Customer Customer"}
        8 {"In Progress Progress"}
        9 {"Waiting Materials Materials"}
        10 {"Dispatched"}
        11 {"Escalate"}
        12 {"Waiting Vendor Vendor"}
        13 {"Waiting Approval Approval"}
        14 {"On Hold Hold"}
        15 {"Scheduled"}
        16 {"Approved by by PRM"}
        17 {"Managed App App Delivery"}
        18 {"Cancelled"}
    }

}