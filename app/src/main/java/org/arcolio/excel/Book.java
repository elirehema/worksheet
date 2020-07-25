package org.arcolio.excel;

/**
 * -- This file created by eli on 15/07/2020 for poixss
 * --
 * -- Licensed to the Apache Software Foundation (ASF) under one
 * -- or more contributor license agreements. See the NOTICE file
 * -- distributed with this work for additional information
 * -- regarding copyright ownership. The ASF licenses this file
 * -- to you under the Apache License, Version 2.0 (the
 * -- "License"); you may not use this file except in compliance
 * -- with the License. You may obtain a copy of the License at
 * --
 * -- http://www.apache.org/licenses/LICENSE-2.0
 * --
 * -- Unless required by applicable law or agreed to in writing,
 * -- software distributed under the License is distributed on an
 * -- "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
 * -- KIND, either express or implied. See the License for the
 * -- specific language governing permissions and limitations
 * -- under the License.
 * --
 **/
public class Book {

    public Book(String title, String author, Integer price, String description, String uriSegment, String theme, String publisher, String mssidn) {
        this.title = title;
        this.author = author;
        this.price = price;
        this.description = description;
        this.uriSegment = uriSegment;
        this.theme = theme;
        this.publisher = publisher;
        this.mssidn = mssidn;
    }

    private String title;

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public String getAuthor() {
        return author;
    }

    public void setAuthor(String author) {
        this.author = author;
    }

    public Integer getPrice() {
        return price;
    }

    public void setPrice(Integer price) {
        this.price = price;
    }

    public String getDescription() {
        return description;
    }

    public void setDescription(String description) {
        this.description = description;
    }

    public String getUriSegment() {
        return uriSegment;
    }

    public void setUriSegment(String uriSegment) {
        this.uriSegment = uriSegment;
    }

    public String getTheme() {
        return theme;
    }

    public void setTheme(String theme) {
        this.theme = theme;
    }

    public String getPublisher() {
        return publisher;
    }

    public void setPublisher(String publisher) {
        this.publisher = publisher;
    }

    public String getMssidn() {
        return mssidn;
    }

    public void setMssidn(String mssidn) {
        this.mssidn = mssidn;
    }

    private String author;
    private Integer price;
    private String description;
    private String uriSegment;
    private String theme;
    private String publisher;
    private String mssidn;
}
