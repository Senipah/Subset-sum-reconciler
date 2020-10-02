# Subuset Sum Reconciler

## Introduction

This is a subset sum problem solver written in Excel VBA.

A common question in Excel is "I have a list of numbers and I want to see which of them add up to a specific total".

This is something most people think should be fairly trivial to achieve in Excel. In reality, however, it ain't all that easy. The question is a variation on a well known problem in computer science called the [Subset sum problem](https://en.wikipedia.org/wiki/Subset_sum_problem).

It can be done with Solver, but there is a variable limit and Solver will only return one possible solution.

As it is something that crops up so often I thought I'd share a workbook I have that can calculate this. [Click here to download it](https://github.com/Senipah/Subuset-sum-reconciler/raw/main/bin/subset_sum_reconciler.xlsm) (xlsm file). This file uses VBA to do the calculation. It uses dynamic programming to offset time complexity with space complexity but given a big list of numbers it still may take too long to be feasible.

## Dependencies

This file depends on (and includes) a copy of my [VBA-Better-Array](https://github.com/Senipah/VBA-Better-Array) library. If you don't want to use this it is possible to alter the code to use VBA's built in Collection or Array types without too much effort.