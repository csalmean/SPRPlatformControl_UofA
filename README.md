# SPR_platform_control
Software suite developed in 2023 to control a surface plasmon resonance measurement setup in Prof. Michael Serpe's lab at the University of Alberta. Modified by Dr. Christopher Salmean (University of Toronto) from basis written by Professor Serpe's team.  

**It is important to credit Professor Serpe's team for writing the initial code, upon which this is based. A rough breakdown of my individual contribution can be found at the end of this readme.**  

Measurement of Surface Plasmon Resonance (SPR) is carried out by directing a laser beam onto a metal-dielectric interface. Momentum matching must be performed using a prism or periodic grating, and in our case the Otto geometry is used (prism-dielectric-metal). Light is shone through a hemicylindrical prism, onto the gold film. At specific incident angles, surface plasmons will become excited and the reflected light will dim significantly. We wish to find the angle and magnitude of this dip.  

![Closeup of setup b](https://github.com/csalmean/SPRPlatformControl_UofA/assets/133036780/ef647971-65ce-4f34-824b-f08369fcdf12)

The experimental setup consists of the sample with controllers for experimental conditions, stepper motors for control of laser/detector arms (Newport XPS-D), and an optical power sensor/meter (Newport 2936). Thanks go to [plasmon360](https://github.com/plasmon360) for use of their [Newport power meter repository](https://github.com/plasmon360/python_newport_1918_powermeter/tree/master).

![Image of Setup](https://github.com/csalmean/SPRPlatformControl_UofA/assets/133036780/d635a574-d724-4351-864f-81500e4062bd)

This software uses a graphical interface perform the following functions:  
1- Homing and movement of optical arms  
2- Queueing of wait periods and experimental runs  
3- Display of queued jobs and estimation of time to execute  
4- Graphical display of experimental data in realtime  
5- Accumulation of experimental data and saving as file with unique filename  

**My individual contributions to the software include:**  
- Rewriting of Newport XPS software to control both optical arms simultaneously rather than sequentially,  
- Rewriting of Newport 2936 communication protocol, removing unnecessary buffer operations, implementation of ring buffer/continuous read operations,  
- Implementation of job queue to reduce burden on experimentalists,  
- Implementation of graphical interface to copy and move queue elements within queue,  
- Alteration of data storage and saving datastructures for greater functionality while saving,  
- Addition of functions to change save directory, check for existing files and rename to avoid overwriting.

These modifications increased the rate of data acquisition by a factor of around 3.5, and reduced the frequency of experimenter interaction by a similar factor.
