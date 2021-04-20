# Common View Specification Schema

> The following is the schema for the Common View Specification

```vb
/*!
 * Licensed Materials - Property of IBM
 * © IBM Corp. 2021. All Rights Reserved.
 * US Government Users Restricted Rights - Use, duplication or
 * disclosure restricted by GSA ADP Schedule Contract with IBM Corp.
 */
    
{
    "$schema": "http://json-schema.org/draft-06/schema#",
    "title": "IBM Planning Analytics Universal View Specification",
    "type": "object",
    "required": ["MDX"],
    "additionalProperties": false,
    "properties": {
        "MDX": {
            "type": "string"
        },
        "Meta": {
            "type": "object",
            "additionalProperties": true,
            "properties": {
                "Aliases": {
                    "type": "object",
                    "additionalProperties": false,
                    "patternProperties": {
                        "^\\[.+\\]\\.\\[.+\\]$": {
                            "type": "string"
                        }
                    }
                },
                "ExpandAboves": {
                    "type": "object",
                    "additionalProperties": false,
                    "patternProperties": {
                        "^\\[.+\\]\\.\\[.+\\]$": {
                            "type": "boolean"
                        }
                    }
                },
                "ContextSets": {
                    "type": "object",
                    "additionalProperties": false,
                    "patternProperties": {
                        "^\\[.+\\]\\.\\[.+\\]$": {
                            "type": "object",
                            "additionalProperties": false,
                            "properties": {
                                "Expression": {
                                    "type": "string"
                                },
                                "SubsetName": {
                                    "type": "string"
                                },
                                "IsPublic": {
                                    "type": "boolean"
                                }
                            },
                            "oneOf": [
                                {
                                    "required": ["Expression"]
                                },
                                {
                                    "required": ["SubsetName"]
                                }
                            ],
                            "dependencies": {
                                "IsPublic": "SubsetName"
                            }
                        }
                    }
                }
            }
        }
    }
}
```

You can use a Common View Specification to embed additional state information in your Exploration View or Quick Report.

A Common View Specification (CVS) is a JSON that can be used to embed additional state information when creating an Exploration View or Quick Report. A CVS is composed of two major parts; the MDX query and a sidecar for additional state information. Data driven mechanisms, such as TurboIntegrator are only concerned with the MDX query, however user interfaces will also consume the sidecar to ensure presentation consistency. By using a CVS, you can generate highly customizable Exploration Views or Quick Reports. For example, using a CVS, you can define aliases and subsets as per the CVS schema input.

To generate an Exploration View or Quick Report from a CVS, use the CreateFromCVS API method. For more information, see [CreateFromCVS (Exploration)](#createfromcvs-exploration) and [CreateFromCVS (Quick Report)](#createfromcvs-quick-report).