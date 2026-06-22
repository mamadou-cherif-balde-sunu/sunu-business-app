from django.db import models
from django.contrib.auth.models import User

class ReponseClient(models.Model):
    CANAL_CHOICES = [('Call Center VIE','Call Center VIE'),('Call Center IARD','Call Center IARD'),('Espace Conseil VIE','Espace Conseil VIE')]
    CSAT_CHOICES = [('Satisfait','Satisfait'),('Insatisfait','Insatisfait')]
    CES_CHOICES = [('Effort faible','Effort faible'),('Effort modere','Effort modere'),('Effort eleve','Effort eleve')]
    STATUT_CHOICES = [('Joint','Joint'),('Injoignable','Injoignable'),('Pas de reponse','Pas de reponse')]
    NPS_CHOICES = [('Promoteur','Promoteur'),('Passif','Passif'),('Detracteur','Detracteur')]
    date_enregistrement = models.DateField()
    cc = models.ForeignKey(User, on_delete=models.CASCADE)
    nom_prenoms = models.CharField(max_length=100)
    contact = models.CharField(max_length=20)
    numero_police = models.CharField(max_length=50, blank=True, null=True)
    objet_visite = models.CharField(max_length=200, blank=True, null=True)
    canal = models.CharField(max_length=30, choices=CANAL_CHOICES)
    date_appel = models.DateField()
    statut_appel = models.CharField(max_length=20, choices=STATUT_CHOICES)
    nps_score = models.IntegerField(blank=True, null=True)
    categorie_nps = models.CharField(max_length=20, choices=NPS_CHOICES, blank=True)
    csat = models.CharField(max_length=20, choices=CSAT_CHOICES, blank=True)
    ces = models.CharField(max_length=20, choices=CES_CHOICES, blank=True)
    motif_insatisfaction = models.TextField(blank=True, null=True)
    observations = models.TextField(blank=True, null=True)
    date_saisie = models.DateTimeField(auto_now_add=True)
    def save(self, *args, **kwargs):
        if self.nps_score is not None:
            if self.nps_score >= 9: self.categorie_nps = 'Promoteur'
            elif self.nps_score >= 7: self.categorie_nps = 'Passif'
            else: self.categorie_nps = 'Detracteur'
        super().save(*args, **kwargs)
    def __str__(self):
        return f"{self.nom_prenoms} - {self.canal} - {self.date_appel}"
    class Meta:
        verbose_name = 'Reponse Client'
        ordering = ['-date_appel']
        
class DemandeModification(models.Model):
    STATUT_CHOICES = [
        ('En attente', 'En attente'),
        ('Approuvee', 'Approuvee'),
        ('Rejetee', 'Rejetee'),
    ]
    saisie        = models.ForeignKey(ReponseClient, on_delete=models.CASCADE)
    demandeur     = models.ForeignKey(User, on_delete=models.CASCADE)
    champ         = models.CharField(max_length=50)
    nouvelle_valeur = models.TextField()
    motif         = models.TextField()
    statut        = models.CharField(max_length=20, choices=STATUT_CHOICES, default='En attente')
    date_demande  = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.demandeur.username} - {self.champ} - {self.statut}"

    class Meta:
        ordering = ['-date_demande']
        verbose_name = 'Demande de modification'