package com.diptam;

import lombok.Data;

@Data
public class Performer {

    private String folio_id;
    private String assignee;
    private String type;
    private String role;
    private String alias;
    private String manager;
    private String manager_alias;
    private String lastupdated;
    private String comment;
}
