RunAction "RunTimeResultFolder", oneIteration
RunAction "TC1_Test_Returns RS02 RS01 P1", oneIteration
RunAction "DataTransfer 1", oneIteration
RunAction "Wait_1 [AT_1.5.1.11 Store returns to DC - (SW31) - RProd]", oneIteration
RunAction "TC2_Test_ReturnsRS02 RS01 P2", oneIteration
