/*
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT License.
 *
 * Code generated by Microsoft (R) AutoRest Code Generator.
 * Changes may cause incorrect behavior and will be lost if the code is regenerated.
 */

import {
  OperationParameter,
  OperationURLParameter,
  OperationQueryParameter,
  QueryCollectionFormat
} from "@azure/core-http";

export const accept: OperationParameter = {
  parameterPath: "accept",
  mapper: {
    defaultValue: "application/json",
    isConstant: true,
    serializedName: "Accept",
    type: {
      name: "String"
    }
  }
};

export const geography: OperationURLParameter = {
  parameterPath: "geography",
  mapper: {
    serializedName: "geography",
    required: true,
    type: {
      name: "String"
    }
  }
};

export const xMsClientId: OperationParameter = {
  parameterPath: "xMsClientId",
  mapper: {
    serializedName: "x-ms-client-id",
    type: {
      name: "String"
    }
  }
};

export const apiVersion: OperationQueryParameter = {
  parameterPath: "apiVersion",
  mapper: {
    defaultValue: "1.0",
    isConstant: true,
    serializedName: "api-version",
    type: {
      name: "String"
    }
  }
};

export const format: OperationURLParameter = {
  parameterPath: "format",
  mapper: {
    serializedName: "format",
    required: true,
    type: {
      name: "String"
    }
  }
};

export const points: OperationQueryParameter = {
  parameterPath: "points",
  mapper: {
    serializedName: "points",
    required: true,
    type: {
      name: "Sequence",
      element: {
        constraints: {
          Pattern: new RegExp(
            "\\s*[-]?\\d+[.]\\d+\\s*[,]\\s*[-]?\\d+[.]\\d+\\s*"
          )
        },
        type: {
          name: "String"
        }
      }
    }
  },
  collectionFormat: QueryCollectionFormat.Pipes,
  skipEncoding: true
};

export const contentType: OperationParameter = {
  parameterPath: ["options", "contentType"],
  mapper: {
    defaultValue: "application/json",
    isConstant: true,
    serializedName: "Content-Type",
    type: {
      name: "String"
    }
  }
};

export const pointsRequestBody: OperationParameter = {
  parameterPath: "pointsRequestBody",
  mapper: {
    serializedName: "pointsRequestBody",
    required: true,
    type: {
      name: "Sequence",
      element: {
        type: {
          name: "Composite",
          className: "CoordinatesPairAbbreviated"
        }
      }
    }
  }
};

export const lines: OperationQueryParameter = {
  parameterPath: "lines",
  mapper: {
    serializedName: "lines",
    required: true,
    type: {
      name: "Sequence",
      element: {
        constraints: {
          Pattern: new RegExp(
            "\\s*[-]?\\d+[.]\\d+\\s*[,]\\s*[-]?\\d+[.]\\d+\\s*"
          )
        },
        type: {
          name: "String"
        }
      }
    }
  },
  collectionFormat: QueryCollectionFormat.Pipes,
  skipEncoding: true
};

export const samples: OperationQueryParameter = {
  parameterPath: ["options", "samples"],
  mapper: {
    constraints: {
      InclusiveMaximum: 2000,
      InclusiveMinimum: 2
    },
    serializedName: "samples",
    type: {
      name: "Number"
    }
  }
};

export const linesRequestBody: OperationParameter = {
  parameterPath: "linesRequestBody",
  mapper: {
    serializedName: "linesRequestBody",
    required: true,
    type: {
      name: "Sequence",
      element: {
        type: {
          name: "Composite",
          className: "CoordinatesPairAbbreviated"
        }
      }
    }
  }
};

export const bounds: OperationQueryParameter = {
  parameterPath: "bounds",
  mapper: {
    serializedName: "bounds",
    required: true,
    type: {
      name: "Sequence",
      element: {
        type: {
          name: "String"
        }
      }
    }
  },
  collectionFormat: QueryCollectionFormat.Csv
};

export const rows: OperationQueryParameter = {
  parameterPath: "rows",
  mapper: {
    constraints: {
      InclusiveMaximum: 1000,
      InclusiveMinimum: 2
    },
    serializedName: "rows",
    required: true,
    type: {
      name: "Number"
    }
  }
};

export const columns: OperationQueryParameter = {
  parameterPath: "columns",
  mapper: {
    constraints: {
      InclusiveMaximum: 1000,
      InclusiveMinimum: 2
    },
    serializedName: "columns",
    required: true,
    type: {
      name: "Number"
    }
  }
};
