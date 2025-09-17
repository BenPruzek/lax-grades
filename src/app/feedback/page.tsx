'use client'

import React, { useState, ChangeEvent, FormEvent } from 'react'
import { Card, CardHeader, CardTitle, CardDescription, CardContent, CardFooter } from "@/components/ui/card"
import { Input } from "@/components/ui/input"
import { Button } from "@/components/ui/button"
import { Label } from "@/components/ui/label"
import { Textarea } from "@/components/ui/textarea"
import { sendFeedback } from '@/lib/actions'
import { FormData } from '@/lib/types'


export default function FeedbackForm() {
    const [formData, setFormData] = useState<FormData>({
        firstName: '',
        lastName: '',
        email: '',
        message: ''
    })
    const [isSubmitting, setIsSubmitting] = useState<boolean>(false)
    const [submitStatus, setSubmitStatus] = useState<'success' | 'error' | 'rateLimit' | null>(null)

    const handleChange = (e: ChangeEvent<HTMLInputElement | HTMLTextAreaElement>) => {
        const { id, value } = e.target
        setFormData(prevData => ({
            ...prevData,
            [id]: value
        }))
    }

    const handleSubmit = async (e: FormEvent<HTMLFormElement>) => {
        e.preventDefault()
        setIsSubmitting(true)
        try {
            await sendFeedback(formData)
            setSubmitStatus('success')
            setFormData({ firstName: '', lastName: '', email: '', message: '' })
        } catch (error) {
            if (error instanceof Error && error.message.includes('Rate limit exceeded')) {
                setSubmitStatus('rateLimit')
            } else {
                setSubmitStatus('error')
            }
        }
        setIsSubmitting(false)
    }

    return (
        <div className="flex flex-col items-center justify-start p-8 gap-8 bg-white dark:bg-transparent min-h-[85vh]">
            <Card className="justify-center items-center bg-[#f6f6ef]">
                <CardHeader>
                    <CardTitle>Send us your feedback</CardTitle>
                    <CardDescription>We&apos;d love to hear your thoughts! Fill out the form below.</CardDescription>
                </CardHeader>
                <form onSubmit={handleSubmit}>
                    <CardContent className="space-y-4">
                        <div className="grid grid-cols-2 gap-4">
                            <div className="space-y-2">
                                <Label htmlFor="firstName">First Name</Label>
                                <Input id="firstName" placeholder="John" value={formData.firstName} onChange={handleChange} required />
                            </div>
                            <div className="space-y-2">
                                <Label htmlFor="lastName">Last Name</Label>
                                <Input id="lastName" placeholder="Doe" value={formData.lastName} onChange={handleChange} required />
                            </div>
                        </div>
                        <div className="space-y-2">
                            <Label htmlFor="email">Email</Label>
                            <Input id="email" type="email" placeholder="m@example.com" value={formData.email} onChange={handleChange} required />
                        </div>
                        <div className="space-y-2">
                            <Label htmlFor="message">Message</Label>
                            <Textarea id="message" placeholder="Let us know your thoughts..." className="min-h-[100px]" value={formData.message} onChange={handleChange} required />
                        </div>
                    </CardContent>
                    <CardFooter className="flex flex-col items-center">
                        <Button type="submit" className="bg-white shadow-1" variant="outline" disabled={isSubmitting}>
                            {isSubmitting ? 'Submitting...' : 'Submit Feedback'}
                        </Button>
                        {submitStatus === 'success' && <p className="text-green-600 mt-2">Feedback submitted successfully!</p>}
                        {submitStatus === 'error' && <p className="text-red-600 mt-2">An error occurred. Please try again.</p>}
                        {submitStatus === 'rateLimit' && <p className="text-red-600 mt-2">Too many submissions. Please try again later.</p>}
                    </CardFooter>
                </form>
            </Card>
        </div>
    )
}