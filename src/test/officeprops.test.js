/* 
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
const OP = require('./../officeprops');
var fs = require("fs");

var pathtofiles="./src/test/files/";

const readFile = (path, opts) =>
    new Promise((res, rej) => {
        fs.readFile(path, opts, (err, data) => {
            if (err) rej(err)
            else res(data)
        })
    })


it('Should parse metadata correctly', () => {
    expect.assertions(4);
    return readFile(pathtofiles+"1testdoc.docx").then((file,err)=>{
        if (err) throw err;
        return OP.getData(file).then(metaData=>{
            expect(metaData.editable.title.value).toEqual("A title");
            expect(metaData.editable.lastModifiedBy.value).toEqual("Torkel Helland Velure");
            expect(metaData.editable.manager.value).toEqual("Torkel Manager");
            expect(metaData.editable.manager.value).not.toEqual("wrong manager");
        });
    })
});

it('Should translate dates correctly', () => {
    expect.assertions(2);
    return readFile(pathtofiles+"1testdoc.docx").then((file,err)=>{
        if (err) throw err;
        return OP.getData(file).then(metaData=>{
            expect(metaData.editable.created.tvalue).toEqual("Fri Mar 16 2018 15:33:00 GMT+0000 (GMT)");
            expect(metaData.editable.modified.tvalue).toEqual("Fri Mar 16 2018 15:36:00 GMT+0000 (GMT)");
        });
    })
});

it('Should translate totalTime correctly', () => {
    expect.assertions(2);
    return readFile(pathtofiles+"1testdoc.docx").then((file,err)=>{
        if (err) throw err;
        return OP.getData(file).then(metaData=>{
            expect(metaData.editable.totalTime.tvalue).toEqual("3 minutes");
            expect(metaData.editable.totalTime.value).toEqual("3");
        });
    })
});



it('1: Should translate ISO8601-time to minutes', () => {
    expect.assertions(2);
    return readFile(pathtofiles+"Doc1.odt").then((file,err)=>{
        if (err) throw err;
        return OP.getData(file).then(metaData=>{
            expect(metaData.editable.totalTime.tvalue).toEqual("1 minute");
            expect(metaData.editable.totalTime.value).toEqual("PT60S");
        });
    })
});

it('2: Should translate ISO8601-time to minutes', () => {
    expect.assertions(2);
    return readFile(pathtofiles+"longertime.odt").then((file,err)=>{
        if (err) throw err;
        return OP.getData(file).then(metaData=>{
            expect(metaData.editable.totalTime.tvalue).toEqual("3 minutes");
            expect(metaData.editable.totalTime.value).toEqual("PT3M43S");
        });
    })
});

it('Should translate docsecurity to correct string', () => {
    expect.assertions(2);
    return readFile(pathtofiles+"1testdoc.docx").then((file,err)=>{
        if (err) throw err;
        return OP.getData(file).then(metaData=>{
            expect(metaData.editable.docSecurity.tvalue).toEqual("None");
            expect(metaData.editable.docSecurity.value).toEqual("0");
        });
    })
});

it('Should create array for duplicate properties', () => {
    expect.assertions(3);
    return readFile(pathtofiles+"multiprops.docx").then((file,err)=>{
        if (err) throw err;
        return OP.getData(file).then(metaData=>{
            expect(metaData.editable.company.value).toBeInstanceOf(Array);
            expect(metaData.editable.company.value[0]).toEqual("University of Manchester");
            expect(metaData.editable.company.value[1]).toEqual("Kilburn");
        });
    })
});

    
it('Should edit metadata properties correctly', () => {
    expect.assertions(2);
    return readFile(pathtofiles+"1testdoc.docx").then((file,err)=>{
        if (err) throw err;
        return OP.getData(file).then(function(metaData){
            metaData.editable.title.value = "something else";
            return OP.editData(file,metaData).then(function(blob){
                    return OP.getData(blob).then(function(metaData){
                        expect(metaData.editable.title.value).toBe("something else");
                        expect(metaData.editable.title.value).not.toBe("not this");
                });
            })
        });
    });
});

it('Should edit metadata properties correctly when duplicate elements are present', () => {
    expect.assertions(2);
    return readFile(pathtofiles+"multiprops.docx").then((file,err)=>{
        if (err) throw err;
        return OP.getData(file).then(function(metaData){
            metaData.editable.company.value = ["Something else", "ABCDE"];
            return OP.editData(file,metaData).then(function(blob){
                    return OP.getData(blob).then(function(meta){
                        expect(meta.editable.company.value[0]).toBe("Something else");
                        expect(meta.editable.company.value[1]).toBe("ABCDE");
                });
            })
        });
    });
});

it('Should edit attributes correctly', () => {
    expect.assertions(4);
    return readFile(pathtofiles+"Doc1.odt").then((file,err)=>{
        if (err) throw err;
        return OP.getData(file).then(metaData=>{
            expect(metaData.editable.paragraphs.value).toBe("0")
            expect(metaData.editable.pages.value).toBe("1")
            metaData.editable.pages.value="27";
            metaData.editable.paragraphs.value="10";
            return OP.editData(file,metaData).then(blob=>{
                return OP.getData(blob).then(meta=>{
                    expect(meta.editable.paragraphs.value).toBe("10")
                    expect(meta.editable.pages.value).toBe("27")
                })
            })
        });
    })
});


it('Should create new textnode for edited empty node', () => {
    expect.assertions(3);
    return readFile(pathtofiles+"2testdoc.docx").then((file,err)=>{
        if (err) throw err;
        return OP.getData(file).then(function(metaData){
            expect(metaData.editable.company.value).toBe("");
            metaData.editable.company.value = "Google";
            return OP.editData(file,metaData).then(function(blob){
                    return OP.getData(blob).then(function(metaData){
                        expect(metaData.editable.company.value).toBe("Google");
                        expect(metaData.editable.company.value).not.toBe("not this");
                });
            })
        });
    });
});

it('Should remove all metadata', () => {
    expect.assertions(2);
    return readFile(pathtofiles+"1testdoc.docx").then((file,err)=>{
        if (err) throw err;
        return OP.removeMetaData(file).then(function(blob){
                return OP.getData(blob).then(function(metaData){
                    expect(Object.keys(metaData.editable).length).toBe(0);
                    expect(Object.keys(metaData.readOnly).length).toBe(0);
            })
        });
    });
});

it('Should parse headingPairsAndParts correctly', () => {
    expect.assertions(1);
    var slideTitles = "Apache Performance Tuning,Agenda,Introduction,Redundancy in Hardware,Server Configuration,Scaling Vertically,Scaling Vertically,Scaling Horizontally,Scaling Horizontally,Load Balancing Schemes,DNS Round-Robin,Example Zone File,Peer-based: NLB,Peer-based: Wackamole,Load Balancing Device,Load Balancing,Linux Virtual Server,Example: mod_proxy_balancer,Apache Configuration,Example: Tomcat, mod_jk,Apache Configuration,Tomcat Configuration,Problem: Session State,Solutions: Session State,Tomcat Session Replication,Session Replication Config,Caching Content,mod_cache Configuration,Make Popular Pages Static,Static Page Substitution,Tuning the Database Tier,Putting it All Together,Monitoring the Farm,Monitoring Solutions,Monitoring Caveats,Conference Roadmap,Current Version,Thank You";
    return readFile(pathtofiles+"1testppt.pptx").then((file,err)=>{
        if (err) throw err;
        return OP.getData(file).then(function(metaData){
            expect(metaData.readOnly.slideTitles.value.join(",")).toBe(slideTitles);
        });
    });
});

it('Should throw error on invalid file', () => {
    expect.assertions(1);
    return readFile(pathtofiles+"invaliddoc.docx").then((file,err)=>{
        if (err) throw err;
        return OP.getData(file).then(function(metaData){
            
        }).catch(e=>{expect(e.message).toBe("Error: File not valid");});
    });
});


it('Should work with xlsb', () => {
    expect.assertions(2);
    return readFile(pathtofiles+"Book1.xlsb").then((file,err)=>{
        if (err) throw err;
        return OP.getData(file).then(metaData=>{
            expect(Object.keys(metaData.editable).length).toBeGreaterThan(0)
            return OP.removeMetaData(file).then(function(officeFile){
                return OP.getData(officeFile).then(function(meta){
                    expect(Object.keys(meta.editable).length).toBe(0)
                })
            })
        });
    })
});

it('Should work with xlsm', () => {
    expect.assertions(2);
    return readFile(pathtofiles+"Book1.xlsm").then((file,err)=>{
        if (err) throw err;
        return OP.getData(file).then(metaData=>{
            expect(Object.keys(metaData.editable).length).toBeGreaterThan(0)
            return OP.removeMetaData(file).then(function(officeFile){
                return OP.getData(officeFile).then(function(meta){
                    expect(Object.keys(meta.editable).length).toBe(0)
                })
            })
        });
    })
});

it('Should work with xlsx', () => {
    expect.assertions(2);
    return readFile(pathtofiles+"Book1.xlsx").then((file,err)=>{
        if (err) throw err;
        return OP.getData(file).then(metaData=>{
            expect(Object.keys(metaData.editable).length).toBeGreaterThan(0)
            return OP.removeMetaData(file).then(function(officeFile){
                return OP.getData(officeFile).then(function(meta){
                    expect(Object.keys(meta.editable).length).toBe(0)
                })
            })
        });
    })
});

it('Should work with docm', () => {
    expect.assertions(2);
    return readFile(pathtofiles+"Doc1.docm").then((file,err)=>{
        if (err) throw err;
        return OP.getData(file).then(metaData=>{
            expect(Object.keys(metaData.editable).length).toBeGreaterThan(0)
            return OP.removeMetaData(file).then(function(officeFile){
                return OP.getData(officeFile).then(function(meta){
                    expect(Object.keys(meta.editable).length).toBe(0)
                })
            })
        });
    })
});


it('Should work with docx', () => {
    expect.assertions(2);
    return readFile(pathtofiles+"Doc1.docx").then((file,err)=>{
        if (err) throw err;
        return OP.getData(file).then(metaData=>{
            expect(Object.keys(metaData.editable).length).toBeGreaterThan(0)
            return OP.removeMetaData(file).then(function(officeFile){
                return OP.getData(officeFile).then(function(meta){
                    expect(Object.keys(meta.editable).length).toBe(0)
                })
            })
        });
    })
});

it('Should work with dotm', () => {
    expect.assertions(2);
    return readFile(pathtofiles+"Doc1.dotm").then((file,err)=>{
        if (err) throw err;
        return OP.getData(file).then(metaData=>{
            expect(Object.keys(metaData.editable).length).toBeGreaterThan(0)
            return OP.removeMetaData(file).then(function(officeFile){
                return OP.getData(officeFile).then(function(meta){
                    expect(Object.keys(meta.editable).length).toBe(0)
                })
            })
        });
    })
});

it('Should work with dotx', () => {
    expect.assertions(2);
    return readFile(pathtofiles+"Doc1.dotx").then((file,err)=>{
        if (err) throw err;
        return OP.getData(file).then(metaData=>{
            expect(Object.keys(metaData.editable).length).toBeGreaterThan(0)
            return OP.removeMetaData(file).then(function(officeFile){
                return OP.getData(officeFile).then(function(meta){
                    expect(Object.keys(meta.editable).length).toBe(0)
                })
            })
        });
    })
});


it('Should work with ppsm', () => {
    expect.assertions(2);
    return readFile(pathtofiles+"pp.ppsm").then((file,err)=>{
        if (err) throw err;
        return OP.getData(file).then(metaData=>{
            expect(Object.keys(metaData.editable).length).toBeGreaterThan(0)
            return OP.removeMetaData(file).then(function(officeFile){
                return OP.getData(officeFile).then(function(meta){
                    expect(Object.keys(meta.editable).length).toBe(0)
                })
            })
        });
    })
});

it('Should work with ppsx', () => {
    expect.assertions(2);
    return readFile(pathtofiles+"pp.ppsx").then((file,err)=>{
        if (err) throw err;
        return OP.getData(file).then(metaData=>{
            expect(Object.keys(metaData.editable).length).toBeGreaterThan(0)
            return OP.removeMetaData(file).then(function(officeFile){
                return OP.getData(officeFile).then(function(meta){
                    expect(Object.keys(meta.editable).length).toBe(0)
                })
            })
        });
    })
});

it('Should work with pptm', () => {
    expect.assertions(2);
    return readFile(pathtofiles+"pp.pptm").then((file,err)=>{
        if (err) throw err;
        return OP.getData(file).then(metaData=>{
            expect(Object.keys(metaData.editable).length).toBeGreaterThan(0)
            return OP.removeMetaData(file).then(function(officeFile){
                return OP.getData(officeFile).then(function(meta){
                    expect(Object.keys(meta.editable).length).toBe(0)
                })
            })
        });
    })
});

it('Should work with potm', () => {
    expect.assertions(2);
    return readFile(pathtofiles+"pp.potm").then((file,err)=>{
        if (err) throw err;
        return OP.getData(file).then(metaData=>{
            expect(Object.keys(metaData.editable).length).toBeGreaterThan(0)
            return OP.removeMetaData(file).then(function(officeFile){
                return OP.getData(officeFile).then(function(meta){
                    expect(Object.keys(meta.editable).length).toBe(0)
                })
            })
        });
    })
});

it('Should work with potx', () => {
    expect.assertions(2);
    return readFile(pathtofiles+"pp.potx").then((file,err)=>{
        if (err) throw err;
        return OP.getData(file).then(metaData=>{
            expect(Object.keys(metaData.editable).length).toBeGreaterThan(0)
            return OP.removeMetaData(file).then(function(officeFile){
                return OP.getData(officeFile).then(function(meta){
                    expect(Object.keys(meta.editable).length).toBe(0)
                })
            })
        });
    })
});

it('Should work with xltm', () => {
    expect.assertions(2);
    return readFile(pathtofiles+"Book1.xltm").then((file,err)=>{
        if (err) throw err;
        return OP.getData(file).then(metaData=>{
            expect(Object.keys(metaData.editable).length).toBeGreaterThan(0)
            return OP.removeMetaData(file).then(function(officeFile){
                return OP.getData(officeFile).then(function(meta){
                    expect(Object.keys(meta.editable).length).toBe(0)
                })
            })
        });
    })
});

it('Should work with xltx', () => {
    expect.assertions(2);
    return readFile(pathtofiles+"Book1.xltx").then((file,err)=>{
        if (err) throw err;
        return OP.getData(file).then(metaData=>{
            expect(Object.keys(metaData.editable).length).toBeGreaterThan(0)
            return OP.removeMetaData(file).then(function(officeFile){
                return OP.getData(officeFile).then(function(meta){
                    expect(Object.keys(meta.editable).length).toBe(0)
                })
            })
        });
    })
});

it('Should work with odt', () => {
    expect.assertions(3);
    return readFile(pathtofiles+"Doc1.odt").then((file,err)=>{
        if (err) throw err;
        return OP.getData(file).then(metaData=>{
            expect(metaData.editable.creator.value).toBe("Torkel Helland Velure")
            metaData.editable.creator.value="Something else";
            return OP.editData(file,metaData).then(blob=>{
                return OP.getData(blob).then(meta=>{
                    expect(meta.editable.creator.value).toBe("Something else")
                    return OP.removeMetaData(blob).then(function(officeFile){
                        return OP.getData(officeFile).then(function(meta){
                            expect(Object.keys(meta.editable).length).toBe(0)
                        })
                    })
                })
            })
        });
    })
});

it('Should work with odp', () => {
    expect.assertions(3);
    return readFile(pathtofiles+"pp.odp").then((file,err)=>{
        if (err) throw err;
        return OP.getData(file).then(metaData=>{
            expect(metaData.editable.creator.value).toBe("Torkel Velure")
            metaData.editable.creator.value="Something else";
            return OP.editData(file,metaData).then(blob=>{
                return OP.getData(blob).then(meta=>{
                    expect(meta.editable.creator.value).toBe("Something else")
                    return OP.removeMetaData(blob).then(function(officeFile){
                        return OP.getData(officeFile).then(function(meta){
                            expect(Object.keys(meta.editable).length).toBe(0)
                        })
                    })
                })
            })
        });
    })
});

it('Should work with ods', () => {
    expect.assertions(3);
    return readFile(pathtofiles+"Book1.ods").then((file,err)=>{
        if (err) throw err;
        return OP.getData(file).then(metaData=>{
            expect(metaData.editable.creator.value).toBe("Torkel Velure")
            metaData.editable.creator.value="Something else";
            return OP.editData(file,metaData).then(blob=>{
                return OP.getData(blob).then(meta=>{
                    expect(meta.editable.creator.value).toBe("Something else")
                    return OP.removeMetaData(blob).then(function(officeFile){
                        return OP.getData(officeFile).then(function(meta){
                            expect(Object.keys(meta.editable).length).toBe(0)
                        })
                    })
                })
            })
        });
    })
});






