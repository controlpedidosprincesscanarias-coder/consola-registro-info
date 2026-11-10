# -*- coding: utf-8 -*-
"""
tarifa_core.py
Recursos compartidos: rutas, logo embebido, configuración persistente,
lógica de negocio de archivos y widgets reutilizables de UI.
Central Compras Princess Canarias
Autor: Victor Alejandro Martin Andrade (VAMA)
(c) VAMA 2026

Símbolos exportados:
    resource_path(relative_path)    — ruta compatible .py / .exe PyInstaller
    LOGO_B64                        — logo Princess en base64
    get_config_path / cargar_config / guardar_config
    TARIFA_NAME / FOLDER_TARIFAS / FOLDER_RECURSOS / FOLDER_PROYECTO
    get_subcarpetas / crear_estructura_proyecto / buscar_tarifa
    listar_xls / renombrar_a_tarifa
    ConsolaLog                      — widget Text con colores y scroll
    CanvasProgressBar               — barra de progreso animada con Canvas
"""

import os
import shutil
import tkinter as tk
from tkinter import ttk
from datetime import datetime
import json
import sys

from tarifa_constants import (
    resource_path,
    C_DARK, C_MID, C_SIDEBAR, C_ACCENT, C_MUTED, C_LOG_BG,
    C_LOG_FG, C_LOG_INFO, C_LOG_WARN, C_LOG_ERR, C_LOG_OK,
)



# ---------------------------------------------------------
#  LOGO PRINCESS (base64 embebido)
# ---------------------------------------------------------
LOGO_B64 = (
    "iVBORw0KGgoAAAANSUhEUgAAAJAAAACQCAIAAABoJHXvAAABCGlDQ1BJQ0MgUHJvZmlsZQAAeJxj"
    "YGA8wQAELAYMDLl5JUVB7k4KEZFRCuwPGBiBEAwSk4sLGHADoKpv1yBqL+viUYcLcKakFicD6Q9A"
    "rFIEtBxopAiQLZIOYWuA2EkQtg2IXV5SUAJkB4DYRSFBzkB2CpCtkY7ETkJiJxcUgdT3ANk2uTml"
    "yQh3M/Ck5oUGA2kOIJZhKGYIYnBncAL5H6IkfxEDg8VXBgbmCQixpJkMDNtbGRgkbiHEVBYwMPC3"
    "MDBsO48QQ4RJQWJRIliIBYiZ0tIYGD4tZ2DgjWRgEL7AwMAVDQsIHG5TALvNnSEfCNMZchhSgSKe"
    "DHkMyQx6QJYRgwGDIYMZAKbWPz9HbOBQAAApwklEQVR42u1deVgT19qfmSQESAwgW2RHRaqiKK64"
    "oVXUutty9Vq11lqXaqtWbWn7ea3VurXaau3jrXUp17qC+1Z3a0UqiqIWwR1QQA2yQwgkM/P98Xv6"
    "PnNDiIjaijfnDx+cnDnnzHnPuy+HYWzN1mzN1mraOI7jOM62D3WssSz70p7Ilwm3GIYJCwvr2LEj"
    "y7I2PHvRm1wuZxgmJibm+vXrDMPIZDIbhr3QTRAEhmE8PDyaNGkSERHB8/zLCrOXim+dPn1aEIRL"
    "ly4pFAqZTPbyMbOXjdADSKGhoXPnzuV53sbJXnQMi4+PF0WxsrLSZDK1a9fuJWZmLwnMLl68KIpi"
    "RUWFKIrHjh2zqWUvLqhYlpXL5RcuXBBFkf+zNW3alCR+Gw97gZooijKZzGQybdmyBRKjKIocxw0d"
    "OtSKTlYX8e/lOXqiKDIM07FjR6mlo1+/fkC46jQBKAO29veYOQIDA00mkyiKgiDwPC+K4rVr1zQa"
    "DQimmTDJMMxrr73Wv3//l4xm1iUR0dvbu6SkBACjf8PDw1mWNZMV8d+9e/fevHmzbgGMe5kAplQq"
    "FQoFPQSSNW/eXBRFMwyDihYcHOzm5mZvby8IwktsL35B9WWGYSIjI8GxRFGkP27cuOHi4iKlivjD"
    "w8OjuLi4srIyMDCwDiHZS4VhXl5ezJ9GRcBAEISgoKBWrVpBaDTrrFarFQqFv78/U3c8Mi8Vs/X1"
    "9a0qOoqi6OjoWBXDxo4diz/wlg1gf4NMHxwcbBH5Ro4cCciZMTz8t0mTJjY97O8BWFXiJpPJeJ4f"
    "Pnx4v379yOHC87xcLic4NWrUiEawAewvbSaTySIsOY4bO3YseBjLsqIourq6hoeHo4Onp6cNYH+P"
    "0HH37t2qWw84hYeH29nZAbcYhomIiLC3tweAAwIC7Ozs6opk/1IBrKioqLqfHB0dXVxcIEPKZLIZ"
    "M2ZI5XtXV1cbD/sbWr169aoDmIuLC2xUPM/36tWrQ4cOgBzDMCqVys3Nra4IilxdQaCa+CEfPnxo"
    "8Tk0syZNmoBaDhw4EIYrCCAsy9YhVaxuAAxmi6o2XLP24MEDK7/+8MMPnp6eLMu6u7vTOFJ9wAaw"
    "Z8acnJ2d27dvD13KCqqVl5db/kiO43ney8srPDxcFEUI9CQ0MgxTUVFh42HPEmBeXl6JiYkHDhzw"
    "9PSEpGeGDUCUR48eVYcoAHazZs08PDxCQ0PNukFaqROS/YsOMEjbN2/ePHPmTL9+/X799ddOnTqZ"
    "TCaL5DEnJ4epxoyL/i1atOjSpQvLshiBftXr9YytPasGzSk6OhrW94qKimnTpnF/NrJokDQBH6ZZ"
    "g28sNzf31KlTUos+noeFhTE2N+azagBG7969eZ43Go3Y4v3790t/BVDnzJkjiqLRaBRr1jAUz/Mh"
    "ISF1BWB1YImQv0+ePJmRkSGXywVBMJlM/fv3P3DgQKNGjaQh2Y9lQiTNS1tRURH0ARsPe8ZgQywN"
    "wtl4nu/Xr19CQoI0jP6xcrlZBBUgVFRUZNFEYgNY7ZUwGN1TUlJolxHU5uHhcfDgwXbt2lVWVnIc"
    "Z9H4a31kqNuVlZUwCtsA9izbt99+K43OkMvlJpPJ0dFxzZo1MplMEIT8/PzaHQiboPAcThbHMQxz"
    "/PhxMzkQf+/du7dRo0aIQrQoJVpskBUzMjLs7e2Zlzpv8+8R7lmW/fDDDwVBMJMDAaHr16+3aNGi"
    "uLiYxL8aSolGo1Fq+7CRxGfJzEpKSqriAfhZkyZNFixYkJubW2tB1NaevY3K19c3Ly8PIdYW6Rtp"
    "xDXHsLKyMh8fHxuGPRdZ8d69ez/++CPcWlWZnCAItdj0iooKWI1tUuJzwbNjx45Vhw21QxFHR0d3"
    "d3fG5l55HkgmimJOTg4w6ekRgmVZQRCUSiWEDhvAnpd8/6x2FlpdWVlZenq6TSF7LvSQZVkXF5ec"
    "nJwnlS8sNugD8fHxjC3y9/nJHQUFBcuWLXuGCPFSVod4sehhgwYNysrKaq4gV9eggE+ZMoX500Fj"
    "w7DnIncUFBRAQX5KJJPJZKIonjlzpg7pznUPYDKZzGAwHD9+/JnssiiK0hxAG8CeV9u0adPT2yag"
    "HrRo0YKxxQc8V1mRYRilUnnr1q2nlBXhvK5bAR1171hBeaqoqEAw4dOzMTs7O5tY/1coZLt370Y4"
    "cK3pIcuy165du3LlCkweNur1HIV7hmHq169fUFBQa+EeMv3kyZPrkExfVzEMZNDJyelpKrUBNYuL"
    "i21a81/RgBMbN258okDEqnapAwcOMLbCKn8ZJztx4sRT0tVmzZrVofTLOgww8K1z5849JQdSq9WO"
    "jo42xfmvABjHcampqadPn7ZSr+2xraKiwmg02ljMX9GoIlvt1Ge8kpyc/Ng8QRuGPUtZ8fr16+Xl"
    "5bVwQPM8bzKZUlNTpVWNbAB7vlRRJpPduXNn/vz5SGzh/7tJq98QjIFbJpNJoVDI5fK65b1kGEbO"
    "1OWGjXZzc5PJZNXpZJSxQrEF6Jmdnb127dpNmzbByQIcffGjBOo2wACJTZs2NW/e3MXFheM4o9H4"
    "6NEjmUzm4eERGBjo6uoqrUCg1+sLCwuvXbu2d+/eLVu26HS6qnyx1vLLX3RGX2KpxM3NrUGDBo6O"
    "jlqtlmXZ/Pz8e/fu5efnS5OL5HK5Vqt1dXU1mUzp6el6vb6upLHU4YbcWfbPJpfL7ezsrChnDg4O"
    "gwYNmjlz5oYNG9LS0kpLS2H4uHPnTlRUFPNi303AvnyQM8sSc3R09PDw8PPzc3FxeeWVV9q0aRMW"
    "FoYKbmasjmXZvLy8V155JS8vj3lRo97kLweoIDgg5t7Z2blr164Mw0RERHTo0MHb29vDw0OlUpkx"
    "P5ijpM1oNLq6ukZGRm7ZsgXJZzaAPResomza1q1bDx8+fODAgc2aNasqnpAQaFaBQCpz8jwPQcbG"
    "xp4XtPDH2LFjT548iStXwJBMJpPRaDSZTICBWWQAHuIP9KysrMSvdShXrE5Cq127drhvBa2yspLg"
    "hEaQwx9WfDFlZWUrVqxQKBQvMrTYugstURRnzZo1Z84ctVoNXOE4riYCnsFgyMvLc3d3LysrW7Ro"
    "0e3btzt06JCTk7N///7bt2+/4GJ9HQaYIAhnzpzp1KmT9HlmZmZ2djbP8+Q0yc3Nzc7OPnv2rL29"
    "fVhY2IULFw4dOvTw4UNPT8/S0lKzin0Y1qY4PxejlCiKjRs3btq0qSAI2dnZ7u7uKpXqyJEjT1Q4"
    "Slrgo07cnfMSWjrMLLlVr6OCDAKQ1zlpsG4DjOy5VL8DwHiJRS32WR1qcgOS0GWTY/9mgFEJT6I5"
    "dJwtgkcul1PQoG2jXywM8/Ly8vHxUSgURqOxpKTk0aNHjx49Iji9+KLXSwgwlmW9vLy0Wq1KpfL2"
    "9m7YsGFQUJCvr6+Li4tMJvP393dyciK0Ky4uzszMPH/+/L59+w4ePGgymWww++sa/BRTpkyxHhRd"
    "WVmJ25PNniclJbVp0+ZZGXtsUbo1ZV2tW7c+evQoLD1SY091OTwIsoAdSKfTeXt7P+WVrs+weECd"
    "bo835IAVPXjw4Oeff3Z2du7UqZMoikA7mUx2/vz59evXnzhxIikp6d69ewzDeHp6wm4kk8k4jqus"
    "rNRoNCqVat++fSiSVzssp+qWNpjVqMEkOnToUIpKR1hS7969zQwH3bp1S05Opg4wHxQVFeESoVrs"
    "OF4JDAzctWtXQkKC2e1ttmaNk73//vtIPgAzKy0t9fHxIZc81U/w8vLS6XRUwgsAHjp0KPOE3ncy"
    "Urz33nsQO0VR7NevH/Nie/GfL4d6ot7e3t5SOpmdnf3gwQOp50kURTs7u5ycnDVr1lAJL2w0bpl/"
    "IsxgWdbe3v7AgQOrVq1ydXWtqKjgeb7WmPq/BTBAKCAgQPrfzMxMiOxS1RhZkUjjITxgWdbZ2flJ"
    "pQxBENzd3fv16wdkVSgUMpns2rVrzP+wR7imAIOwQDfoYr9QosmiaTUjI6O0tFTqW3rS+02AQ1qt"
    "FsEXOBkpKSkXLlywWH7PBrD/2jtRFFUqldmFrChzbbHp9XqDwSB9kpWVVQuANWnShGXZyspKOzs7"
    "o9E4efJkVMD+39Wyar53Pj4+Hh4eUuUsOTm5OuqkVCqRnU+E8Y8//qgFKWvZsiXLskql8vTp0336"
    "9Dl9+vTzMJogOlgul0N0sqIvchwn/7NZPzfUk0a20p8WgJhzuVz+VHYG7Hj//v0h8oHoFRYWWqwL"
    "iZlatGghremalZWFQDN48avKeNKNwLcBTpmZmYWFhe+//z6Glb5Irzz2+61va3Udqn5XDTfRyqRV"
    "YWZlARank9ccwxA7BpGPZdnMzExc/1R1Glz8hFu7kGewf//+srIypVJJnIzYG5Zl0dvr6uq6YsWK"
    "mJgYKkhPrAtB8BZRDT9JmZzFMA2cCYzg5OTUrFmzoKAgtVqdk5Nz9erVmzdvSl+hmPvGjRv36NGj"
    "YcOG6enpGzduNAvtxpjoGRgY2L17944dOzo5OZWWliYnJ2/fvv3hw4fUH0oLOnt5eUVERHTq1Mnd"
    "3b28vDwlJWX37t21DzDBKd6wYQOUMOhV27dvt6gP4cnq1avRGdvapk0bPK9Xr15kZGSPHj2Y/74o"
    "MTQ0dPr06cuXL1+6dGl0dHT9+vXpp+Dg4IULFy5fvhxXXNJJb9y4cXR09Ny5c6lYFOltMpmsb9++"
    "y5Yti4uL++CDD6qeVvpv586dN2zYkJ2dLbWuGQyGpKQkqpCDlTRu3Hj79u3l5eXUbf/+/VLaRX9E"
    "Rkbu2rWrpKTEzGin0+lat26NRVLnDh06bNy4EfUrpK20tLRPnz611DiBYefPnwdJhIXwiy++YKrk"
    "F+OIqVSqe/fuwSIsiuKuXbvw6/Dhw+/cuYMFzZw5Ew/bt2+/b98+qVkyNzcXtn+VSvXdd99RYFrD"
    "hg1poujoaNqR6dOnM5Kbz4cMGSINfBNF8a233pJ+OTbL09MTR5CM10ifJRvpiBEjWJYFJx49ejQ0"
    "d7M4xoiICOkZ8vb2jo2NpTFRrsdkMsEynpubi6BHbJqTk9OPP/5IeWwVFRUGgwGdDQaDXq9HfNET"
    "MzNAq379+nl5eWTbFUVx8ODBVeGPwz569GjCRb1eD3V7/vz5tDWCINy9e1ej0SxevBijmUwmg8GA"
    "Kohvv/02zM1k4jIajbdu3QJIXFxctm7diqHKy8tNJtPOnTsJfdesWUNZe9hWnufv3bunUqlwmPD9"
    "zZs3R6kqszBF+rrx48fT5yxatMhi1Qjcwc4wDIDaqlWru3fvWrSGY5bXXnuNxgwICLh8+bLFzjgK"
    "EydOrCV64Qvbtm1LhkFprXcpF8XBUSqVaWlpPM8bDAZRFMeNG9esWTOkjhOFBMCAsnhIK758+bJM"
    "JsNNYFIKfOjQIYZhHBwc/vjjD5xH2ty4uDjIk9gCrJDuBhNF8ebNm9gmoIJWq01PTxdFUa/Xi6J4"
    "9+7dzz//fNasWadOnULs8EcffURgWLZsGVZy48aNt99++9KlSxgW47dr1w7b6uXllZWVhYUVFxdP"
    "nz69ffv2b775ZmJiItawevVqMsmqVCqMYzAYKioq5syZ0759+6FDhx47dgzDwuxQy+oIeG3MmDHS"
    "zU1MTKQkH0hrgJyTk9P27duxblEUv/nmG4Zhbt26de7cuW3bthG8ya+GbpWVlXFxcVOnTn3vvfcC"
    "AwOdnJzS0tJmzJiB2/TQZ+HChQzDODo69u/ff9GiRYAWTu7ixYuHDBkC7Jw3b15YWNisWbMQWooO"
    "u3fvlgpve/fuJWjt378fti6GYdq0aSOK4vLlyxmGwW0sEydOxPIKCwtxc+bnn3+OjRYEISUlhcKE"
    "161bR2MuXbqUdi88PFwQhF9//dXR0RGCK8Mw//rXv6hzbGwsdfb396+srLx+/bqHh0ftDdyYAweN"
    "XFzYPumICoVi9OjRaWlphCtffvklEA4GLQI5kSyALTY2tmXLltIZ7ezsMDJIIkgEbMegitOnTycb"
    "tCAIN27cAJeOjIzECBqNBllfeBcYg3dHjBhBh2D9+vU0o1wuDwoKGjt2LDLMOI4LDg4uKyvDCP/3"
    "f/8HBN25cycAJoriwIEDsQk+Pj56vZ6o2ahRo+zs7JRKpVwud3BwaNq0qUajIVqlVquzsrIEQcAa"
    "Pv30U7lcbm9vDxt6kyZNntZYimmOHTtGhFsQhDVr1oSGhmo0Gmdn55YtW06bNi0pKYlAdfbs2b59"
    "+zKS+sccxy1cuJAABjQtKSkZOXIkHQs0clQGBgZCJMNGIKML5kRpxSJAXa/Xd+/eHZjBcVyrVq2g"
    "L6IR4dJoNLB/opg2yIMZnyDZdceOHfjk4uJiLy8vmUzm5uaWl5eHeVeuXEmHADybpJV169ZhqdJN"
    "p2G7d+9OwguRejMlr/bQwpsk9dG1oQDMw4cPyesB6W7r1q0DBgzAW0SCwT82b95MaMHzfFlZWc+e"
    "PdHNTBACgYXvDWc2LS0No4FQSP1t+Ox3332XEIVhmPHjx9M6b968qVQqsVmfffYZ8cW2bduacXXg"
    "Fp60atWK53nMvm3bNnzRggUL8KVxcXEgsPi0uXPnkmSBVc2ZM4eUHDqCWNvEiROJVmPxq1atog9/"
    "Wsc6yVRSs8WtW7eGDRt26dIlGAwzMzM3btw4YsQILy8vM21MKuvTLmOVs2fPJsZukQivXLkScqAo"
    "ips3bybA+/n5gfqT0HHkyBH6FfPu2rWLOA22g+M4V1dXOIPozlOLMhhmX7VqFZG+f/zjHwzDvPrq"
    "q1hMbGwsKAcALCXR0pIta9euhXGHDi7+eOONN6SXnKHznj17QAmfthAg3h8+fLgU5Q8fPowtcHNz"
    "8/DwkNZqwrmTnhFSUCAUYH35+fmurq7VnSZQj5SUFMKSqVOnEv0x83pTDVia183NDXoo+kCYZhjm"
    "008/pRPwxhtv0HZXpSiOjo6gKLDAKRSKtm3bQqtZu3YtiVr0dQ0bNiwvL5cWBwHwLl++3L59e9pG"
    "uijh/v37UpgBjzMzM7HUp/LNYqYlS5aYSRwk8FQHJzNUGzRoEOlGoihu3bq1upVhC0JCQsCEsAXh"
    "4eEEMFysSOI+SYC02iFDhtCW5eTkwD6iUqkyMjLA0nQ6HZxzFq8iYximR48etI9xcXFdunQBsk6b"
    "No2pEg6EBb/77rvSqn60gIqKCqmphe6bxncRUtJb8+bNY54m4gjTHD58WIphw4YNI07zWOkTm/jd"
    "d99JQT5mzBiLB5z6g8hgy8jwgc+AFYMc3F27diV+TlYxQRDKy8sFQQD/ZximT58+NODOnTut00Po"
    "+Oh8+/ZtKGHIm7Z4LjHUmDFjYHyh3SdD+YcffkjdsKUDBgx48OCBtA8xi2+//fapLFIODg5Q4InJ"
    "wwpcQ5MJqMeVK1foSyoqKiDyWblS6tChQ0QPUaIeHxAYGAi+gqESExPpxNBqMzIy6PD2798fwy5f"
    "vhxQFEXxk08+qY5bmN20iU9etWoV0NRKhXuMFhISgndJ4wSdNBgMYPBS6cPPzw+CqFTZx7Jbt25d"
    "G6OUVOIg6pSdnQ1eWhOcxQhNmjSBOQq7nJSUVB1e4qGHh0dRURGd8UWLFpEm+9Zbb0lxHfYbirlj"
    "GKZnz54Ezvv372OjWZaFVQUDWjSq0ewajUan09EgMFRaOfL0IdRh1qxZiKmVWltg76ZTQp3feeed"
    "kpISacCSIAijRo2q9kg9FsNCQ0PhM6TqaWVlZTWsngaAde7cWaFQYN0Mwxw5cgRRi9UxvIiICI1G"
    "gxRYhmESExPJ8wnVGGXBcnNz4TGQhixCBsM5PXz4cElJCcSQoKAgGt/iPZkkowcFBbm7u1OJt6NH"
    "j8KUU11QAvWknN2lS5eOHz8elSioD3YfgQ7UWS6Xr1+/PioqCoeDPtOKrPh4gEFfoQwihAXUEFvx"
    "Cg6XlCNW53rGQ3hK6daOS5cuMQxTWVnp4OAARoJ24MCBvLw85Mgg6EOpVPbt2xcsjWXZvXv3QiEN"
    "CAhwcnKiHLKqFS3hl6qsrIQGhg3FYvz8/Myyp6SqFcdxDg4OVIGWgoX+85//nDp1CsOCnKCuFdKx"
    "qbPJZLKzszt8+PC2bdukna3cLMM9NvAG9JQQH2EBNeRePM8rFArIeGBm2dnZFy5cYCxdm4L+Dg4O"
    "3bt3p+lu376dnZ2N0xoaGurv7y89+9SNLNSBgYGYtKCgID4+HjsI/kE+Uq1WSwcOso8gCD169ACB"
    "JTuZWdwR+cbIccqybK9evVJTUydMmEBiM6DLcdyePXsI+2/evJmamhoSEnL16tU5c+ZIra+gFvBA"
    "4Ujl5+ejNK5F9yxnZbsFQYArlnQj6BZMzW6pwYKCg4MbNmxIhzQhIaG0tFRKLszoZ6tWrQAVPDx/"
    "/rzRaMSJhmXEZDLhHmAYw2jviGACgRITE3U6XcOGDcPCwojCYNihQ4cKggBLGNhGjx49Tpw4AfOu"
    "Wq2Wrh+2R57nAVee59u1a7dr16533nkHFpaAgIA2bdqArVJogiAISLUGPs2bN89oNI4dO9bPzy88"
    "PJwEQqhDPM+jM+jk4sWLc3NzLW7R49lPu3btqBIJQjPwPTWROLBNUFDIsm7lsi6pJZvUrOnTp2/Y"
    "sAF25Pj4eBIcbt26Jb2CQyrdQYycPXv24MGDy8vL165d26lTJ9AfIFlpaSnoHsMw7u7un3zyCe5a"
    "x5P58+eb2ZnI+t66deuYmBgcvrlz56rV6tLSUp7nHz58uGLFChoTXouEhAQM8ssvv+AcZGZmCoKQ"
    "n5//008/derUiTZBqVTu2LED35ucnKxUKmujh0mNcrR9x44dq5rjbX0EfCFZ1q2kHlW9vN5oNMJ5"
    "P3jw4Pr160Moxy6AhkhrAGg0Gmg2gA28UykpKQ0aNGjdurWZn0yn0y1btuz777+HReP48eNqtRoL"
    "7tu3r7SIMP44ceLEtm3bcBSKi4unTp3KsiyQTGpkunjx4pYtWzZt2kSO9czMTPhvoQhKPX9Xr17d"
    "tm3bhg0brl69SlaVkJAQpnapWVg9HLhw3QqCEBMT80T2LrIwYZfv3bsHO1bV40NWJViTyVVhNBrh"
    "xYiKiuJ5Hn50o9EIUyRWIo3Tkl58//XXX0MD8fDwKCgoINORWZbbxo0bYUOBqOLo6Hjnzp2qdia0"
    "gwcPIuacYZh9+/YR4lossJOYmAjplGGY77//nrqRxC9taWlpHTp0qL1pCjt47tw5omYYtGnTpjVB"
    "WGxiQEAAjDpQgWGltYJejRs3lhYYysrK6ty5M379+eefpZ8HKEoNHAMGDKDzfvr06Y4dO1a1rqGw"
    "EY1vNBr/9a9/SaN3pGQc3Whnr1+/PmbMGDqIfn5+0KylYxKM79+//9lnn9E5UKvVyA5B2IhZ54KC"
    "gqVLl8KaU0tokU2TNgjFV0VRTEhIqAkPw8SRkZHSXYbToToExSvwwuAs+/n5SQMrDhw4cOLEiePH"
    "jy9evNjFxYW0bxyg5s2bFxUVCYIwe/ZsMihQHIdKpfrll1/MUkNfffVVi7ZBmUwGDzLaH3/8MX36"
    "dDghSWfw8PA4c+ZM1UCPs2fPzpgxQ6vV0mjAWjJqSFtycvLs2bMhiNYQWqyV2OwGDRpERUU9evQo"
    "PT09Ozu7oqLC09OzqKjo7t27NQR8vXr1fHx8VCqVvb29Xq9PS0vDJfLVnRJBEIKDg+fNm7d9+3ZE"
    "atSwCC8W3LRpU3t7e5Sip5A/aVxi165dQ0ND7ezsUlJSjh8/zvO8xfHRv1evXsHBwVeuXElISEAf"
    "6owOLMt27969ffv2Xl5eLMveuHHj5MmTV69eJcwGDtHsHTt2DA8P9/X1VSgUmZmZJ0+evHjxIn5C"
    "qmPdzvAwM19ZDCGxbiKyHnJrnbeb9a86oxUaY315VTs/kYjBPjbinJHUSnnScj9Sy2wNC64QO6lF"
    "fgoFEVsh1FT95bEnGp2t9MRSpZCwXq3KzNIvrbn5bABmJcjZ7KFZGRwrZ4qilC0+r25GM4R70kNT"
    "dZbqVijtYPGjzH6t7qfqBqxuQx473bNsZiToeSSzPqsxrWem1HBGM03U+kXrNZyx5h8os342IZia"
    "mT6VSiVRCUgKLMs2btzY1dW1sLAQ9w4RKec4DsF70kh6Ozs7YkiQyoj34javqsIIx3GNGzf29fXl"
    "OK60tLSGuKVQKCjWg0gQ7S9u8yAGaWZfd3BwCAoK8vT05HmeZCXCKnd3d9iFy8rKzMJ4aEwpK6FC"
    "w7BdAZ8I2IIgqNVqOAqMRqNZat0ThGenp6cjEo2iMEeMGJGenu7q6ip1eCOQFqoxSr/b29tfvHix"
    "oKCAhEydTldQUDB//nyGYU6fPp2Tk/Po0aP8/Py8vLyCggKoTZMmTbp9+zY5sejfcePGwSSBTf/9"
    "99+p8JX1MxsdHZ2RkaHT6fLz83NycrZu3Qq1lyJt7t27l5ubm5+fj+XB/mtnZ/fll18WFxfTjIjb"
    "xVuIFSd9cf369Wq1GocDprXMzEydTpeXl5ednf3zzz8HBwfDgJKfn//o0SOdTnfnzh1MWlBQEBoa"
    "qlAoVq5ciZgXcMEFCxZUa8CzzsMDAgJcXFykW+Ps7BwQEEDsvWvXrnv37j148ODIkSN5nv/000/j"
    "4uJ69ep14sSJZcuWqVQqxC39/vvvMTExDg4OSHht0aLFhQsXNm3aZGdnB4snany4uLhIMx4gRkdF"
    "Ra1duzY2Nnb58uV6vT44OLhFixYKhUKagwt8NUsxgmHe399/4sSJgiB4enq+++678fHxbdu2zczM"
    "hJ4uk8mio6PhrpPJZPC9ffnllx999NHChQt37NgBJ4C9vT1wSyaT7dy5s1u3buPHj7948WJ4eDgS"
    "5uEUZRjG19fXzc1t6tSpMplMq9W+9957vXv3DgsLS01NnT17tsFgCA0NnTp16hdffJGVlaVQKG7d"
    "uvX9999PmDBh9uzZBw4ckMvl4eHhQOgnyFskQxHP85MmTaJAT1gXeZ5HKh/DMBcuXECmM7W0tDTS"
    "RdB0Ot2SJUukm5ufnw9UM5tx5syZJpOJMAzHIiEh4c6dO1YYu0U8o4Bl+DDJJltRUfHVV1/hv4cP"
    "Hz5z5ozZiw4ODmVlZQgTqjogInwQhowGswjSWBBRgtOAptVqKaYPDdZ62ARwRikstUaS8GMFZekt"
    "GVJaDF9RWFjYv//9b5icwfBWrVrVrFkzHF4wCZlMplKpELdMpEylUikUCgcHB7AZCgSTsmhwFHd3"
    "d1zVgHhQNDOpslu3boMGDUIYQdV4W0RNK5XKoqKia9euIeeHmLRCoaB/OY7DJYvAeAqIJw878o72"
    "7NlD6bCbN2/W6/Vw6kvZGD7/wYMHGRkZyNTGJjg7O3McV79+ffQBAYMtQjpdLQEG/5M04An3JWDE"
    "li1bIjAbDl+j0ciybGJioiiKYWFhMMWiP/6A0RPDYsDy8nIauTq9CkRs8ODBqNMhVdGwj9HR0QMG"
    "DFCr1d9++61ZgLRUDQAJ1Wq12dnZUmXIaDTCxIeMmPz8/NTUVCTdUMQq3WWAquvSgjF6vT41NRV2"
    "WynG46oJiAI5OTkUikKfgNdzcnLu378/depUaYCsFWLIPdbHr1arnZyc3Nzc3NzcnJyc6tWrR5vr"
    "7+/PsixWQyEJyAr19/e37st2cnLSarW+vr6o6mfxTEG++uKLL9LT03fv3r1kyRLEepBXRRAEe3v7"
    "iRMnrlu3bvPmzf7+/hhKKiXD5gvv+4IFCzw8PNauXUsaj1Kp1Gq13t7eDRo08PLycnBw4Hn+gw8+"
    "UKvVycnJkyZNwg6CADg4OGi1WlA8qRnz7t27Xl5e5J8z/dnkcvmKFSs0Gk1MTIzFy9gRAzF58mQf"
    "H5/Lly+PGjUK0HpiNYY4dlFRkdFo1Ov1SCbU6/VIDkD9B8TSghwTnfTw8KDUFUycl5f33XffkRnX"
    "zs4OwX4UWqTT6dzc3BiG+fjjjwVBqColent7wyicnZ0NKZRMDCqVqqioqFu3bl26dIGMYGakN5lM"
    "sbGxhw8fvn37dnJyMgy++BUGWSAuVoJoBoZh2rRpA9vuuXPnKMQB/Gbx4sU0Av5dt26dwWDAsleu"
    "XKnX67du3Xr06NH09PTz58+TwwG70bt3b1EUW7RoId20rl27ImPs5MmTcHxXRxWtYRiqv69ZsyYq"
    "KmrYsGHDhg2Lior64YcfcAyJx5gNjX20YljieV6j0Wzbti0yMrJ37949evQYMmQIVKuqZxDxEdnZ"
    "2W+++Wbfvn1LS0vj4uKWLVsG5U+hUJSVlZ07d27ZsmWRkZHTpk0bNWrU6tWrx40bJ018hucC6mB+"
    "fj4B29HR8fr165GRkT179uzZs2ePHj0gLsnl8gsXLnTu3HnChAlNmzZNSEgYO3Ys0ZWq0raZ6Zbj"
    "OI1GA/QVRTEvL88KxgCfTp8+3bZt2xkzZnTp0iUpKWnIkCGkzj6ZlEgJrNSQWtOgQQOoTSgYwPx3"
    "rjiF81nEMI7jCgsLkWFmpjYhEU+KYWZWTQcHB6AaRYh2794dcXPIQouIiEhLS9NqtZjrm2++IcUW"
    "hvyHDx/a2dlhtKNHj549e7Y6aYuqk1y6dMloNPr5+YGCrV+/nsKWofnu27fvwYMH5PIlb4Zarb5x"
    "40Z6ejpGQwczDDMzlwQFBd2+fbuoqMjd3d1i9ObjYahSqSBoQUiDDxenCdK2n5+fVIz09fVFhDNT"
    "fRkVnD4MayUuXyoHIhaqvLx8ypQpFRUV//znPxFXGh0d/c4772RmZiJzMiQk5JNPPnnw4IE0Lgpp"
    "SAaDYc6cOR4eHq+88gpRCBLnpMsAv0Ghs6ysrKlTp8rl8t69e/M8n52dHRwcTME/6BYUFIS0M4yJ"
    "oZRKZWlp6dy5cwMCAuBZrO4boUFiups3b3700UcajaZz584Wr799PMAgIFGT3tYEsgsvJaRzURRf"
    "ffVVlmWTkpIYq3Vv2CpNyn6k+X3YAkQ4IX5Nr9fb29vDQLBw4cKsrKyPP/548ODB48aNq1+//pEj"
    "R6AIS4kwQpqQ8oRsl+pWQoAEQigUCmRRQP6Oj49v1aqVg4MDVAKO49zd3YOCgk6dOmVG9iEzI2AN"
    "8TnVkTiz6W7dugVLlUUV8zEAo3hK6egENp1Ot2/fvgkTJri6uhoMBoPBoNFopkyZcuzYsfv375O2"
    "VLUCiiAIJSUlqBxANWlp8IKCAirpBxIvLV3btWtXFxeXo0eP+vv716tXD/QnNjY2JiZm7dq1e/bs"
    "gaqAqUkcx+A5OTk6nQ6hsXhClRZI6wC1xH/h/n/99dc5joMRZM2aNY6OjlOmTEGACc/zM2fO5Dhu"
    "w4YN0vNNukRmZmZxcTFCj6QbSLtadbrhw4dzHIfw2apMXf5YfxjVv6CYLGki7IwZM86dO3f58uWv"
    "vvrKZDLNmjVLoVAgqIhecXZ2luaQwRz85ptvgtMYjUaFQhETE/Pbb79Bw/3pp5+go8jl8tjY2Ndf"
    "f93Nze3EiRN6vT4kJOTDDz9MSUnZtGkTKjnMnz9/0qRJSOLPyMjYsWPH5MmTNRpNWVnZoUOH1Go1"
    "xRni0r1r166hYAk2q23btj/88INCocDNzseOHYuLi0tISLhy5UpSUhLP8926dRs9evT27dvj4+Nx"
    "7/OPP/749ddfN2rUKCkpqWvXrmPGjPnmm29SUlLs7OwqKytVKhVymUAYULCiT58+kKiZP7N+ASfU"
    "Fzx69Gh8fPyVK1cQVBEVFbV69eqUlBSLVbWsBfgrFIqwsLBDhw7duHGDaIWbm5u7u3tsbKzBYOA4"
    "Li8vb//+/U2bNh0zZkzPnj2Tk5NHjhwpnYxl2dDQ0N9///38+fN4KJfLGzVqJJPJUG3M3d1dq9We"
    "O3fu+vXrnp6eTk5OXl5eDRo0wPPU1NTc3NyBAwcOGjSoX79+/v7+P/300/jx4wsLC8vKyu7cufPW"
    "W2/1798/JCQkNjZ2yZIlERER77//vtFojImJqaio8PHxQToa+AFELxcXlx07dgiC4O/vz3Gcj4+P"
    "VqvFdDk5Ob/99lvHjh179uw5ePDgvn37qlSqr776CvoGPmf//v16vX7EiBGDBg1yd3efN2/evHnz"
    "iJz4+fmVlZVBwQANVCgUKpWKYjo0Go2/v//OnTsLCwvBLDt16hQZGTlo0KDIyEiO4+bNmzd37tza"
    "5IdZFwTMvOzSOl1SYm2BbdaqTplCoYDlqaovG64DaoTN1ieyIgIQLTGjLmYdKN+XnliJDHgsAKT1"
    "72rvca4aHGIW32JmKafLQ80GMXtYNVCObhyqWi7TzLFLwS1mUTEUGAM0ohh3Myu+NIyg6jKwTulX"
    "gy9IuQ6e0JWmVZUwKzPSBkpfkQYCVX39L4qZ+cvGr/r8WS3jsV/0bL/XVirc1mzN1mzN1mzN1l6k"
    "9v/FR53Zmux1+wAAAABJRU5ErkJggg=="
)


# ---------------------------------------------------------
#  CONFIGURACION PERSISTENTE
# ---------------------------------------------------------
def get_config_path():
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, "tarifa_config.json")


def cargar_config():
    path = get_config_path()
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def guardar_config(config: dict):
    path = get_config_path()
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
    except Exception:
        pass


# ---------------------------------------------------------
#  LOGICA DE NEGOCIO
# ---------------------------------------------------------
TARIFA_NAME      = "TARIFA.xls"
FOLDER_TARIFAS   = "TARIFAS"
FOLDER_RECURSOS  = "RECURSOS"
FOLDER_PROYECTO  = "PROYECTO"


def get_subcarpetas(carpeta_proyecto):
    """Devuelve las rutas de TARIFAS y RECURSOS dentro del proyecto."""
    return (
        os.path.join(carpeta_proyecto, FOLDER_TARIFAS),
        os.path.join(carpeta_proyecto, FOLDER_RECURSOS),
    )


def crear_estructura_proyecto(escritorio):
    """Crea PROYECTO/TARIFAS y PROYECTO/RECURSOS en el escritorio."""
    base = os.path.join(escritorio, FOLDER_PROYECTO)
    tarifas  = os.path.join(base, FOLDER_TARIFAS)
    recursos = os.path.join(base, FOLDER_RECURSOS)
    os.makedirs(tarifas,  exist_ok=True)
    os.makedirs(recursos, exist_ok=True)
    return base


def buscar_tarifa(carpeta_proyecto):
    """Busca TARIFA.xls dentro de la subcarpeta TARIFAS del proyecto."""
    ruta = os.path.join(carpeta_proyecto, FOLDER_TARIFAS, TARIFA_NAME)
    return ruta if os.path.isfile(ruta) else None


def listar_xls(carpeta_tarifas):
    archivos = []
    try:
        for entry in os.scandir(carpeta_tarifas):
            if entry.is_file() and entry.name.lower().endswith(".xls"):
                ts = entry.stat().st_mtime
                fecha = datetime.fromtimestamp(ts).strftime("%d-%m-%Y")
                archivos.append({"nombre": entry.name, "ruta": entry.path, "fecha_mod": fecha})
    except PermissionError:
        pass
    archivos.sort(key=lambda x: x["nombre"].lower())
    return archivos


def renombrar_a_tarifa(ruta_origen, carpeta_tarifas):
    destino = os.path.join(carpeta_tarifas, TARIFA_NAME)
    if os.path.isfile(destino):
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup = os.path.join(carpeta_tarifas, f"TARIFA_backup_{ts}.xls")
        shutil.copy2(destino, backup)
    shutil.copy2(ruta_origen, destino)
    # Eliminar el archivo original si es distinto al destino
    if os.path.abspath(ruta_origen) != os.path.abspath(destino):
        try:
            os.remove(ruta_origen)
        except OSError:
            pass
    return destino


# ---------------------------------------------------------
#  BOILERPLATE COMPARTIDO DE VENTANAS Toplevel (punto 6 del roadmap
#  de limpieza, 2026-09-24) -- extraído de las clases VentanaXxxRINFO
#  de crear_registro_info.py y las VentanaXxx de tarifa_dialogs.py,
#  que repetían este mismo código byte a byte. Alcance deliberadamente
#  conservador: solo se extrajo lo que era IDÉNTICO en varias clases;
#  las geometrías con fórmulas de ancho/alto propias (una por caso de
#  uso) se dejaron tal cual, para no forzar una firma común que metiera
#  riesgo real de romper la disposición visual de alguna ventana.
# ---------------------------------------------------------
def cargar_icono_ventana(win):
    """
    Carga el icono estándar (assets/logo.ico) en una ventana Toplevel.
    Idéntico en 13 de las 15 ventanas del proyecto; se ignora en
    silencio si el icono no está disponible (mismo comportamiento que
    tenían todas las copias sueltas).
    """
    try:
        win.iconbitmap(resource_path("assets/logo.ico"))
    except Exception:
        pass


def centrar_sobre_padre(win, parent):
    """
    Centra una ventana Toplevel sobre el centro de su ventana padre
    (no sobre la pantalla completa). Usa el tamaño ya calculado de
    `win` (llamar después de construir la UI y de update_idletasks()
    si hace falta que winfo_width()/height() sean fiables).

    Idéntico en 4 de las ventanas de tarifa_dialogs.py (Selección XLS,
    Abrir tarifa edición, Crear proyecto, Selección tarifa RI).
    """
    win.update_idletasks()
    pw = parent.winfo_rootx() + parent.winfo_width()  // 2
    ph = parent.winfo_rooty() + parent.winfo_height() // 2
    w, h = win.winfo_width(), win.winfo_height()
    win.geometry(f"+{pw - w // 2}+{ph - h // 2}")


def ajustar_geometria_pantalla(win, ancho, margen_alto=20):
    """
    Centra una ventana Toplevel en la pantalla completa, con un ancho
    fijo (`ancho`) y alto ajustado al contenido real más un margen.

    Idéntico salvo el ancho en 3 ventanas de crear_registro_info.py
    (Teléfono, Fechas y Centros del flujo "crear RINFO desde cero").
    """
    win.update_idletasks()
    w = ancho
    h = win.winfo_reqheight() + margen_alto
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    x = (sw - w) // 2
    y = max(20, (sh - h) // 2)
    win.geometry(f"{w}x{h}+{x}+{y}")


# ---------------------------------------------------------
#  WIDGET: CONSOLA DE LOG
# ---------------------------------------------------------
class ConsolaLog(tk.Frame):

    COLORES = {
        "INFO":    C_LOG_INFO,
        "OK":      C_LOG_OK,
        "WARN":    C_LOG_WARN,
        "ERROR":   C_LOG_ERR,
        "SISTEMA": C_ACCENT,
        ">>":      C_LOG_FG,
    }

    def __init__(self, parent, **kwargs):
        super().__init__(parent, bg=C_LOG_BG, **kwargs)
        self._build()

    def _build(self):
        # Detectar mejor fuente monoespaciada disponible en el sistema
        try:
            import tkinter.font as _tkf
            _fams = _tkf.families()
            _mono = next((f for f in ["Cascadia Code", "Cascadia Mono",
                                       "JetBrains Mono", "Consolas", "Courier New"]
                          if f in _fams), "Courier New")
        except Exception:
            _mono = "Consolas"
        _FONT_LOG = (_mono, 8)

        # ── Cabecera estilo macOS ─────────────────────────────────────
        hdr = tk.Frame(self, bg=C_DARK, pady=5)
        hdr.pack(fill="x")

        # Tres puntos de color (close/minimise/maximise)
        _dot_cnv = tk.Canvas(hdr, bg=C_DARK, width=56, height=14,
                              highlightthickness=0)
        _dot_cnv.pack(side="left", padx=(10, 0))
        for _i, _col in enumerate(["#FF5F57", "#FEBC2E", "#28C840"]):
            _cx = 8 + _i * 18
            _dot_cnv.create_oval(_cx - 5, 2, _cx + 5, 12,
                                  fill=_col, outline="")

        tk.Label(hdr, text="VENTANA INMEDIATA  —  Flujo de operaciones",
                 bg=C_DARK, fg=C_ACCENT,
                 font=("Segoe UI", 8, "bold")).pack(side="left", padx=12)

        btn_limpiar = tk.Label(hdr, text="⟳ limpiar",
                                bg=C_DARK, fg=C_MUTED,
                                font=("Segoe UI", 8), cursor="hand2")
        btn_limpiar.pack(side="right", padx=10)
        btn_limpiar.bind("<Button-1>", lambda e: self.limpiar())
        btn_limpiar.bind("<Enter>",
                          lambda e: btn_limpiar.config(fg=C_ACCENT))
        btn_limpiar.bind("<Leave>",
                          lambda e: btn_limpiar.config(fg=C_MUTED))

        btn_guardar = tk.Label(hdr, text="💾 guardar log",
                                bg=C_DARK, fg=C_MUTED,
                                font=("Segoe UI", 8), cursor="hand2")
        btn_guardar.pack(side="right", padx=(0, 4))
        btn_guardar.bind("<Button-1>", lambda e: self.guardar_log())
        btn_guardar.bind("<Enter>",
                          lambda e: btn_guardar.config(fg=C_ACCENT))
        btn_guardar.bind("<Leave>",
                          lambda e: btn_guardar.config(fg=C_MUTED))

        tk.Frame(self, bg=C_ACCENT, height=1).pack(fill="x")

        # ── Área de texto ─────────────────────────────────────────────
        frm = tk.Frame(self, bg=C_LOG_BG)
        frm.pack(fill="both", expand=True)

        _sty = ttk.Style()
        _sty.configure("Thin.Vertical.TScrollbar",
                        background=C_MID, troughcolor=C_LOG_BG,
                        arrowcolor=C_LOG_BG, bordercolor=C_LOG_BG,
                        arrowsize=0, relief="flat", width=6)
        sb = ttk.Scrollbar(frm, orient="vertical",
                            style="Thin.Vertical.TScrollbar")
        sb.pack(side="right", fill="y")

        self.txt = tk.Text(frm, bg=C_LOG_BG, fg=C_LOG_FG,
                           font=_FONT_LOG, wrap="none",
                           state="disabled", cursor="arrow",
                           yscrollcommand=sb.set,
                           insertbackground=C_LOG_FG,
                           relief="flat", padx=8, pady=4,
                           selectbackground=C_MID, bd=0)
        self.txt.pack(side="left", fill="both", expand=True)
        sb.config(command=self.txt.yview)

        _sty.configure("Thin.Horizontal.TScrollbar",
                        background=C_MID, troughcolor=C_LOG_BG,
                        arrowcolor=C_LOG_BG, bordercolor=C_LOG_BG,
                        arrowsize=0, relief="flat", width=4)
        sbx = ttk.Scrollbar(self, orient="horizontal",
                             style="Thin.Horizontal.TScrollbar")
        sbx.pack(fill="x")
        self.txt.config(xscrollcommand=sbx.set)
        sbx.config(command=self.txt.xview)

        for nivel, color in self.COLORES.items():
            self.txt.tag_configure(nivel, foreground=color)
        self.txt.tag_configure("ts",  foreground="#2A4A6A")
        self.txt.tag_configure("sep", foreground="#1E3A5A")

    def log(self, mensaje, nivel=">>", ts=None):
        """
        ts: timestamp ya formado ("%H:%M:%S"), opcional. Si se omite, se
        genera aquí con datetime.now() -- pero eso captura el instante en
        que este método se EJECUTA, no en el que se LLAMÓ. Desde que
        App.log_consola() encola las llamadas desde hilos de fondo con
        self.after(0, ...) (2026-09-24, fix de thread-safety), ambos
        instantes pueden diferir si hay cola pendiente -- App.log_consola()
        ya captura y pasa el ts correcto por este motivo; el valor por
        defecto queda solo para otros llamadores directos.
        """
        if ts is None:
            ts = datetime.now().strftime("%H:%M:%S")
        self.txt.config(state="normal")
        self.txt.insert("end", f" {ts} ", "ts")
        self.txt.insert("end", f"[{nivel:<7}]", nivel)
        self.txt.insert("end", f"  {mensaje}\n", nivel)
        self.txt.config(state="disabled")
        self.txt.see("end")

    def log_batch(self, lineas_niveles, ts=None):
        """
        Inserta una lista de (mensaje, nivel) en un ÚNICO ciclo state=normal/disabled.
        Evita miles de repaints individuales cuando hay logging masivo.
        lineas_niveles: iterable de (str, str)
        ts: timestamp ya formado, opcional -- mismo motivo que en log().
        """
        if not lineas_niveles:
            return
        if ts is None:
            ts = datetime.now().strftime("%H:%M:%S")
        self.txt.config(state="normal")
        for mensaje, nivel in lineas_niveles:
            self.txt.insert("end", f" {ts} ", "ts")
            self.txt.insert("end", f"[{nivel:<7}]", nivel)
            self.txt.insert("end", f"  {mensaje}\n", nivel)
        self.txt.config(state="disabled")
        self.txt.see("end")

    def separador(self, texto=""):
        if not texto:
            linea = "─" * 58
        else:
            pad = max(0, 54 - len(texto))
            linea = f"── {texto} " + "─" * pad
        self.txt.config(state="normal")
        self.txt.insert("end", "\n " + linea + "\n", "sep")
        self.txt.config(state="disabled")
        self.txt.see("end")

    def limpiar(self):
        self.txt.config(state="normal")
        self.txt.delete("1.0", "end")
        self.txt.config(state="disabled")

    def guardar_log(self):
        """Exporta el contenido actual de la consola a un fichero .txt elegido por el usuario."""
        from tkinter import filedialog, messagebox
        contenido = self.txt.get("1.0", "end").strip()
        if not contenido:
            messagebox.showinfo("Log vacío", "No hay entradas en el log para guardar.")
            return
        ts_nombre = datetime.now().strftime("%Y%m%d_%H%M%S")
        ruta = filedialog.asksaveasfilename(
            title="Guardar log",
            initialfile=f"log_{ts_nombre}.txt",
            defaultextension=".txt",
            filetypes=[("Texto plano", "*.txt"), ("Todos los archivos", "*.*")],
        )
        if not ruta:
            return
        try:
            with open(ruta, "w", encoding="utf-8") as f:
                f.write(contenido)
            messagebox.showinfo("Log guardado", f"Log exportado correctamente a:\n{ruta}")
        except Exception as exc:
            messagebox.showerror("Error al guardar", f"No se pudo guardar el log:\n{exc}")


# ---------------------------------------------------------
#  WIDGET: BARRA DE PROGRESO — Canvas custom (animación suave)
# ---------------------------------------------------------
class CanvasProgressBar(tk.Frame):
    """Barra de progreso dibujada con Canvas. Animación suave en modo indeterminado."""
    _TRACK_BG    = "#0D1F35"
    _FILL_CLR    = C_ACCENT
    _DONE_CLR    = "#27AE60"
    _BLOCK_RATIO = 0.32
    _STEP_PX     = 8
    _FPS_MS      = 25

    def __init__(self, parent, **kwargs):
        super().__init__(parent, bg=C_SIDEBAR, **kwargs)
        self._anim_pos = 0
        self._anim_job = None
        self._mode     = "determinate"
        self._value    = 0
        self._maximum  = 100
        self._build()

    def _build(self):
        self._lbl_estado = tk.Label(self, text="Listo",
                                     bg=C_SIDEBAR, fg=C_MUTED,
                                     font=("Segoe UI", 8), anchor="w")
        self._lbl_estado.pack(fill="x")
        frm = tk.Frame(self, bg=C_SIDEBAR)
        frm.pack(fill="x", pady=(2, 0))
        self._canvas = tk.Canvas(frm, height=8, bg=C_SIDEBAR,
                                  highlightthickness=0, bd=0)
        self._canvas.pack(side="left", fill="x", expand=True)
        self._canvas.bind("<Configure>", lambda e: self._redraw())
        self._lbl_pct = tk.Label(frm, text="", bg=C_SIDEBAR,
                                  fg=C_ACCENT, font=("Segoe UI", 8, "bold"),
                                  width=5, anchor="e")
        self._lbl_pct.pack(side="right", padx=(4, 0))

    def _redraw(self):
        c = self._canvas
        c.delete("all")
        w, h = c.winfo_width(), c.winfo_height()
        if w < 4:
            return
        c.create_rectangle(0, 1, w, h - 1, fill=self._TRACK_BG, outline="")
        if self._mode == "determinate":
            ratio = max(0.0, min(1.0, self._value / self._maximum)) if self._maximum else 0
            fw = int(w * ratio)
            if fw > 0:
                c.create_rectangle(0, 1, fw, h - 1,
                                   fill=self._FILL_CLR, outline="")
        elif self._mode == "indeterminate":
            bw  = max(int(w * self._BLOCK_RATIO), 30)
            pos = int(self._anim_pos) % (w + bw) - bw
            x0  = max(0, pos)
            x1  = min(w, pos + bw)
            if x1 > x0:
                c.create_rectangle(x0, 1, x1, h - 1,
                                   fill=self._FILL_CLR, outline="")
        elif self._mode == "complete":
            c.create_rectangle(0, 1, w, h - 1, fill=self._DONE_CLR, outline="")

    def _animate(self):
        if self._mode != "indeterminate":
            return
        w  = self._canvas.winfo_width() or 200
        bw = max(int(w * self._BLOCK_RATIO), 30)
        self._anim_pos = (self._anim_pos + self._STEP_PX) % (w + bw)
        self._redraw()
        self._anim_job = self.after(self._FPS_MS, self._animate)

    def _stop(self):
        if self._anim_job:
            self.after_cancel(self._anim_job)
            self._anim_job = None

    # ── API pública (misma que BarraProgreso) ─────────────────────
    def set(self, valor, maximo=100, estado=""):
        self._stop()
        self._mode    = "determinate"
        self._value   = valor
        self._maximum = maximo
        pct = int(valor / maximo * 100) if maximo else 0
        self._lbl_pct.config(text=f"{pct}%", fg=C_ACCENT)
        if estado:
            self._lbl_estado.config(text=estado)
        self._redraw()

    def indeterminate(self, estado="Procesando..."):
        self._stop()
        self._mode     = "indeterminate"
        self._anim_pos = 0
        self._lbl_estado.config(text=estado)
        self._lbl_pct.config(text="·  ·  ·", fg=C_MUTED)
        self._animate()

    def reset(self, estado="Listo"):
        self._stop()
        self._mode  = "determinate"
        self._value = 0
        self._lbl_pct.config(text="", fg=C_ACCENT)
        self._lbl_estado.config(text=estado)
        self._redraw()

    def completado(self, estado="Completado"):
        self._stop()
        self._mode = "complete"
        self._lbl_pct.config(text="✔", fg=self._DONE_CLR)
        self._lbl_estado.config(text=estado, fg=self._DONE_CLR)
        self._redraw()
        # Restaurar color texto estado tras un momento
        self.after(3000, lambda: self._lbl_estado.config(fg=C_MUTED))

