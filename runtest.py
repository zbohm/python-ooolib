#!/usr/bin/env python3.9
from ooolib.calc import Calc


calc = Calc()
calc.load("test-02.ods")
calc.save("test-02-out-04.ods")
