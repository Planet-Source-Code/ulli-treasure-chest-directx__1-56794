Remember the old days when the Black Avenger of the Spanish Main (alias Tom Sawyer) roamed the Carribean Seas and spread terror among his foes? Well, this is what went overboard one day and you are the lucky one to find it.

Found this on the Internet and modified/debugged it. It will try to adjust the frame rate to your monitors refresh frequency (if your CPU is fast enough). On my Athlon 1800 it uses virtually no resources and still achieves 75 fps. Try modifying line 145 in mD3D to

          D3DWindow.SwapEffect = D3DSWAPEFFECT_FLIP

to achieve the full frame rate (should be around 250 fps).

Tested with NVIDIA GeForce2 MX/MX 400 and Windows XP only !